using System;
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
                    return;null) return;
                
                // Process both tasks in parallelount ?? 0;
                var tasks = new List<Task<bool>> 
                {f (numComments == _prevNumComments) 
                    AICommentReplyTask(),
                    UserMentionTask()
                }; Process both tasks in parallel
                var tasks = new List<Task<bool>> 
                var results = await Task.WhenAll(tasks);
                numComments += results.Count(r => r);
                    UserMentionTask()
                _prevNumComments = numComments;
            }
            catch (Exception ex)ait Task.WhenAll(tasks);
            {   numComments += results.Count(r => r);
                CommonUtils.DisplayError(ex);
            }   _prevNumComments = numComments;
        }   }
            catch (Exception ex)
        private static async Task<bool> AICommentReplyTask()
        {       CommonUtils.DisplayError(ex);
            var app = Globals.ThisAddIn?.Application;
            if (app == null) return false;
            
            var doc = app.ActiveDocument;
            if (doc == null) return false;
            
            var comments = GetUnansweredAIComments(doc.Comments);
            
            if (_isDraftingComment || !comments.Any())Task()
                return false;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            _isDraftingComment = true;edAIComments(doc.Comments);
            try
            {f (_isDraftingComment || !comments.Any())
                var tasks = comments.Select(async comment => 
                {
                    var chatHistory = new List<ChatMessage>
                    {
                        new UserChatMessage($@"{Forge.CultureHelper.GetLocalizedString("[Review] chatHistory #1")}\n""{CommonUtils.SubstringTokens(comment.Range.Text, (int)(ThisAddIn.ContextLength * 0.2))}"""),
                        new UserChatMessage(Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #2"))
                    };
                    var chatHistory = new List<ChatMessage>
                    chatHistory.AddRange(GetCommentMessages(comment));
                    chatHistory.Add(new UserChatMessage(@$"{Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #3")}:\n""{comment.Scope.Text}"""));gth * 0.2))}"""),
                        new UserChatMessage(Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #2"))
                    await AddComment(
                        comment.Replies,
                        comment.Range,ge(GetCommentMessages(comment));
                        RAGControl.AskQuestion(tMessage(@$"{Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #3")}:\n""{comment.Scope.Text}"""));
                            Forge.CommentSystemPrompt,
                            chatHistory,
                            doc.Range(),
                            0.5f,ange,
                            docrol.AskQuestion(
                        )   Forge.CommentSystemPrompt,
                    );      chatHistory,
                            doc.Range(),
                    return true;,
                });         doc
                        )
                var results = await Task.WhenAll(tasks);
                return results.Any(r => r);
            }       return true;
            catch (OperationCanceledException ex)
            {
                CommonUtils.DisplayWarning(ex);l(tasks);
                return false;s.Any(r => r);
            }
            finallyOperationCanceledException ex)
            {
                _isDraftingComment = false;ex);
            }   return false;
        }   }
            finally
        private static async Task<bool> UserMentionTask()
        {       _isDraftingComment = false;
            var app = Globals.ThisAddIn?.Application;
            if (app == null) return false;
            
            var doc = app.ActiveDocument;
            if (doc == null) return false;
            
            var comments = GetUnansweredMentionedComments(doc.Comments);
            foreach (var comment in comments)
            {te static async Task<bool> UserMentionTask()
                List<ChatMessage> chatHistory = new List<ChatMessage>();
                chatHistory.AddRange(GetCommentMessagesWithoutMention(comment));ication.ActiveDocument.Comments);
                chatHistory.Add(new UserChatMessage(@$"{Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #3")}:\n""{comment.Scope.Text}"""));
            foreach (var comment in comments)
                try
                {ist<ChatMessage> chatHistory = new List<ChatMessage>();
                    if (_isDraftingComment) return false;thoutMention(comment));
                    _isDraftingComment = true;ssage(@$"{Forge.CultureHelper.GetLocalizedString("(CommentHandler.cs) [AICommentReplyTask] UserChatMessage #3")}:\n""{comment.Scope.Text}"""));
                    
                    await AddComment(
                        comment.Replies,
                        comment.Range,ment) return false;
                        RAGControl.AskQuestion(
                            new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(CommentHandler.cs) [AIUserMentionTask] UserMentionSystemPrompt"]),
                            chatHistory,
                            Globals.ThisAddIn.Application.ActiveDocument.Range(),
                            0.5f,ange,
                            docrol.AskQuestion(
                        )   new SystemChatMessage(ThisAddIn.SystemPromptLocalization["(CommentHandler.cs) [AIUserMentionTask] UserMentionSystemPrompt"]),
                    );      chatHistory,
                            Globals.ThisAddIn.Application.ActiveDocument.Range(),
                    return true;
                }
                catch (OperationCanceledException ex)       )
                {
                    CommonUtils.DisplayWarning(ex);   
                }
                finally   return true;
                {   }
                    _isDraftingComment = false;erationCanceledException ex)
                }       {
            }                    CommonUtils.DisplayWarning(ex);
            string modelName = $"@{ThisAddIn.Model}";
ment parentComment)        }   }
            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {        }
                new UserChatMessage(GetCleanedCommentText(parentComment, modelName))
            };sagesWithoutMention(Comment parentComment)

            Comments childrenComments = parentComment.Replies;}";
            for (int i = 1; i <= childrenComments.Count; i++)
            {atHistory = new List<ChatMessage>()
                var comment = childrenComments[i];
                string cleanText = GetCleanedCommentText(comment, modelName);w UserChatMessage(GetCleanedCommentText(parentComment, modelName))
                chatHistory.Add(;
                    (i % 2 == 1) ? new AssistantChatMessage(cleanText) : new UserChatMessage(cleanText)
                );mments = parentComment.Replies;
            }   for (int i = 1; i <= childrenComments.Count; i++)
            {
            return chatHistory;
        }       string cleanText = GetCleanedCommentText(comment, modelName);

        private static string GetCleanedCommentText(Comment c, string modelName)
        {       );
            string commentText = c.Range.Text;            }
            return commentText.Contains(modelName) ? commentText.Remove(commentText.IndexOf(modelName), modelName.Length).TrimStart() : commentText;
        }   return chatHistory;

        private static IEnumerable<ChatMessage> GetCommentMessages(Comment parentComment)
        {, string modelName)
            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {string commentText = c.Range.Text;
                new UserChatMessage(parentComment.Range.Text)xt.Remove(commentText.IndexOf(modelName), modelName.Length).TrimStart() : commentText;
            };
            
            Comments childrenComments = parentComment.Replies;tCommentMessages(Comment parentComment)
            for (int i = 1; i <= childrenComments.Count; i++)
            {
                var comment = childrenComments[i];
                chatHistory.Add(   new UserChatMessage(parentComment.Range.Text)
                    (i % 2 == 1) ? new AssistantChatMessage(comment.Range.Text) : new UserChatMessage(comment.Range.Text)            };
                );
            }   Comments childrenComments = parentComment.Replies;
            for (int i = 1; i <= childrenComments.Count; i++)
            return chatHistory;
        }       var comment = childrenComments[i];

        private static IEnumerable<Comment> GetUnansweredMentionedComments(Comments allComments)ntChatMessage(comment.Range.Text) : new UserChatMessage(comment.Range.Text)
        {   );
            List<Comment> comments = new List<Comment>();
            foreach (Comment c in allComments)
            {
                if (c.Ancestor == null &&
                    (c.Range.Text.Contains($"@{ThisAddIn.Model}") ? 
                        (c.Replies.Count == 0 || c.Replies[c.Replies.Count].Author != ThisAddIn.Model) : omment> GetUnansweredMentionedComments(Comments allComments)
                        AreRepliesUnbalanced(c.Replies)))
                {ist<Comment> comments = new List<Comment>();
                    comments.Add(c); c in allComments)
                }   {
            }                if (c.Ancestor == null &&
            return comments;? 
        }               (c.Replies.Count == 0 || c.Replies[c.Replies.Count].Author != ThisAddIn.Model) : 

        private static bool AreRepliesUnbalanced(Comments replies)
        {
            int userMentionCount = GetCommentMentionCount($"@{ThisAddIn.Model}", replies);       }
            int aiAnswerCount = GetCommentAuthorCount(ThisAddIn.Model, replies);            }
            return userMentionCount > aiAnswerCount;
        }

        private static int GetCommentMentionCount(string mention, Comments comments)ents replies)
        {
            int count = 0;
            for (int i = 1; i <= comments.Count; i++)iAnswerCount = GetCommentAuthorCount(ThisAddIn.Model, replies);
            {onCount > aiAnswerCount;
                if (comments[i].Range.Text != null && comments[i].Range.Text.Contains(mention))
                {
                    count++;t GetCommentMentionCount(string mention, Comments comments)
                }
            }            int count = 0;
            return count;
        }   {
ts[i].Range.Text != null && comments[i].Range.Text.Contains(mention))
        private static int GetCommentAuthorCount(string author, Comments comments)
        {       count++;
            int count = 0;
            for (int i = 1; i <= comments.Count; i++)
            {
                if (comments[i].Author == author)
                {
                    count++;t GetCommentAuthorCount(string author, Comments comments)
                }
            }            int count = 0;
            return count;
        }   {

        private static IEnumerable<Comment> GetUnansweredAIComments(Comments allComments)
        {       count++;
            List<Comment> comments = new List<Comment>();
            foreach (Comment c in allComments)
            {
                if (c.Ancestor == null &&
                    c.Author == ThisAddIn.Model &&
                    c.Replies.Count > 0 &&omment> GetUnansweredAIComments(Comments allComments)
                    c.Replies[c.Replies.Count].Author != ThisAddIn.Model)
                {ist<Comment> comments = new List<Comment>();
                    comments.Add(c); c in allComments)
                }   {
            }                if (c.Ancestor == null &&
            return comments;
        }           c.Replies.Count > 0 &&
dIn.Model)
        public static async Task AddComment(Comments comments, Range range, AsyncCollectionResult<StreamingChatCompletionUpdate> streamingContent)
        {
            Word.Comment c = comments.Add(range, string.Empty);                }
            c.Author = ThisAddIn.Model;
            Word.Range commentRange = c.Range.Duplicate;            return comments;

            StringBuilder comment = new StringBuilder();
s comments, Range range, AsyncCollectionResult<StreamingChatCompletionUpdate> streamingContent)
            await Task.Run(async () =>
            {Comment c = comments.Add(range, string.Empty);
                Forge.CancelButtonVisibility(true);
                trye commentRange = c.Range.Duplicate;
                {
                    await foreach (var update in streamingContent.WithCancellation(ThisAddIn.CancellationTokenSource.Token))= new StringBuilder();
                    {
                        if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)n(async () =>
                            break;
                        foreach (var content in update.ContentUpdate)
                        {
                            commentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            commentRange.Text = content.Text; foreach (var update in streamingContent.WithCancellation(ThisAddIn.CancellationTokenSource.Token))
                            commentRange = c.Range.Duplicate;
                            comment.Append(content.Text);       if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                        }     break;
                    }       foreach (var content in update.ContentUpdate)
                }
                finally           commentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                {
                    Forge.CancelButtonVisibility(false);             commentRange = c.Range.Duplicate;
                }                   comment.Append(content.Text);
                c.Range.Text = WordMarkdown.RemoveMarkdownSyntax(comment.ToString());                   }
            });                   }
        }                }



}    }                finally
                {
                    Forge.CancelButtonVisibility(false);
                }
                c.Range.Text = WordMarkdown.RemoveMarkdownSyntax(comment.ToString());
            });
        }
    }
}
