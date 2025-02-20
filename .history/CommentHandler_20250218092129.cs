﻿using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using OpenAI.Chat;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    internal class CommentHandler
    {
        private static int _prevNumComments = 0;
        private static bool _isDraftingComment = false;

        public static async Task Document_CommentsEventHandlerAsync(Word.Selection selection)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                int numComments = doc.Comments.Count;
                
                if (numComments == _prevNumComments) 
                    return;

                // Process both tasks in parallel
                var tasks = new List<Task<bool>> 
                {
                    AICommentReplyTask(),
                    UserMentionTask()
                };

                var results = await Task.WhenAll(tasks);
                numComments += results.Count(r => r);

                _prevNumComments = numComments;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static async Task<bool> AICommentReplyTask()
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            var comments = GetUnansweredAIComments(doc.Comments);
            
            if (_isDraftingComment || !comments.Any())
                return false;

            _isDraftingComment = true;
            try
            {
                var tasks = comments.Select(async comment => 
                {
                    var chatHistory = new List<ChatMessage>
                    {
                        new UserChatMessage($@"{Forge.CultureHelper.GetLocalizedString("[Review] chatHistory #1")}\n""{CommonUtils.SubstringTokens(comment.Range.Text, (int)(ThisAddIn.ContextLength * 0.2))}"""),
                        new UserChatMessage(Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #2"))
                    };
                    
                    chatHistory.AddRange(GetCommentMessages(comment));
                    chatHistory.Add(new UserChatMessage(@$"{Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #3")}:\n""{comment.Scope.Text}"""));

                    await AddComment(
                        comment.Replies,
                        comment.Range,
                        RAGControl.AskQuestion(
                            Forge.CommentSystemPrompt,
                            chatHistory,
                            doc.Range(),
                            0.5f,
                            doc
                        )
                    );
                    
                    return true;
                });

                var results = await Task.WhenAll(tasks);
                return results.Any(r => r);
            }
            catch (OperationCanceledException ex)
            {
                CommonUtils.DisplayWarning(ex);
                return false;
            }
            finally
            {
                _isDraftingComment = false;
            }
        }

        private static async Task<bool> UserMentionTask()
        {
            var comments = GetUnansweredMentionedComments(Globals.ThisAddIn.Application.ActiveDocument.Comments);
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (var comment in comments)
            {
                List<ChatMessage> chatHistory = new List<ChatMessage>();
                chatHistory.AddRange(GetCommentMessagesWithoutMention(comment));
                chatHistory.Add(new UserChatMessage(@$"{Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #3")}:\n""{comment.Scope.Text}"""));

                try
                {
                    if (_isDraftingComment) return false;
                    _isDraftingComment = true;
                    
                    await AddComment(
                        comment.Replies,
                        comment.Range,
                        RAGControl.AskQuestion(
                            new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(CommentHandler.cs) [AIUserMentionTask] UserMentionSystemPrompt"]),
                            chatHistory,
                            Globals.ThisAddIn.Application.ActiveDocument.Range(),
                            0.5f,
                            doc
                        )
                    );
                    
                    _isDraftingComment = false;
                    return true;
                }
                catch (OperationCanceledException ex)
                {
                    CommonUtils.DisplayWarning(ex);
                }
            }
            return false;
        }

        private static IEnumerable<ChatMessage> GetCommentMessagesWithoutMention(Comment parentComment)
        {
            string modelName = $"@{ThisAddIn.Model}";

            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {
                new UserChatMessage(GetCleanedCommentText(parentComment, modelName))
            };

            Comments childrenComments = parentComment.Replies;
            for (int i = 1; i <= childrenComments.Count; i++)
            {
                var comment = childrenComments[i];
                string cleanText = GetCleanedCommentText(comment, modelName);
                chatHistory.Add(
                    (i % 2 == 1) ? new AssistantChatMessage(cleanText) : new UserChatMessage(cleanText)
                );
            }

            return chatHistory;
        }

        private static string GetCleanedCommentText(Comment c, string modelName)
        {
            string commentText = c.Range.Text;
            return commentText.Contains(modelName) ? commentText.Remove(commentText.IndexOf(modelName), modelName.Length).TrimStart() : commentText;
        }

        private static IEnumerable<ChatMessage> GetCommentMessages(Comment parentComment)
        {
            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {
                new UserChatMessage(parentComment.Range.Text)
            };
            
            Comments childrenComments = parentComment.Replies;
            for (int i = 1; i <= childrenComments.Count; i++)
            {
                var comment = childrenComments[i];
                chatHistory.Add(
                    (i % 2 == 1) ? new AssistantChatMessage(comment.Range.Text) : new UserChatMessage(comment.Range.Text)
                );
            }

            return chatHistory;
        }

        private static IEnumerable<Comment> GetUnansweredMentionedComments(Comments allComments)
        {
            List<Comment> comments = new List<Comment>();
            foreach (Comment c in allComments)
            {
                if (c.Ancestor == null &&
                    (c.Range.Text.Contains($"@{ThisAddIn.Model}") ? 
                        (c.Replies.Count == 0 || c.Replies[c.Replies.Count].Author != ThisAddIn.Model) : 
                        AreRepliesUnbalanced(c.Replies)))
                {
                    comments.Add(c);
                }
            }
            return comments;
        }

        private static bool AreRepliesUnbalanced(Comments replies)
        {
            int userMentionCount = GetCommentMentionCount($"@{ThisAddIn.Model}", replies);
            int aiAnswerCount = GetCommentAuthorCount(ThisAddIn.Model, replies);
            return userMentionCount > aiAnswerCount;
        }

        private static int GetCommentMentionCount(string mention, Comments comments)
        {
            int count = 0;
            for (int i = 1; i <= comments.Count; i++)
            {
                if (comments[i].Range.Text != null && comments[i].Range.Text.Contains(mention))
                {
                    count++;
                }
            }
            return count;
        }

        private static int GetCommentAuthorCount(string author, Comments comments)
        {
            int count = 0;
            for (int i = 1; i <= comments.Count; i++)
            {
                if (comments[i].Author == author)
                {
                    count++;
                }
            }
            return count;
        }

        private static IEnumerable<Comment> GetUnansweredAIComments(Comments allComments)
        {
            List<Comment> comments = new List<Comment>();
            foreach (Comment c in allComments)
            {
                if (c.Ancestor == null &&
                    c.Author == ThisAddIn.Model &&
                    c.Replies.Count > 0 &&
                    c.Replies[c.Replies.Count].Author != ThisAddIn.Model)
                {
                    comments.Add(c);
                }
            }
            return comments;
        }

        public static async Task AddComment(Comments comments, Range range, AsyncCollectionResult<StreamingChatCompletionUpdate> streamingContent)
        {
            if (comments == null)
                throw new ArgumentNullException(nameof(comments));

            Word.Comment c = comments.Add(range, string.Empty);
            c.Author = ThisAddIn.Model;
            Word.Range commentRange = c.Range.Duplicate;

            StringBuilder comment = new StringBuilder();

            await Task.Run(async () =>
            {
                Forge.CancelButtonVisibility(true);
                try
                {
                    await foreach (var update in streamingContent.WithCancellation(ThisAddIn.CancellationTokenSource.Token))
                    {
                        if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                            break;
                        foreach (var content in update.ContentUpdate)
                        {
                            commentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            commentRange.Text = content.Text;
                            commentRange = c.Range.Duplicate;
                            comment.Append(content.Text);
                        }
                    }
                }
                finally
                {
                    Forge.CancelButtonVisibility(false);
                }
                c.Range.Text = WordMarkdown.RemoveMarkdownSyntax(comment.ToString());
            });
        }
    }
}
