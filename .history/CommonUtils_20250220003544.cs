﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Resources;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    internal class CommonUtils
    {
        private class RememberAccessChoice
        {
            private Dictionary<string, bool> _rememberAccess = new Dictionary<string, bool>();
            
            public bool IsGrantedAccess(string website)
            {
                bool result;
                _rememberAccess.TryGetValue(website, out result);
                return result;
            }

            public void Grant(string website)
            {
                _rememberAccess[website] = true;
            }

            public void Revoke(string website)
            {
                _rememberAccess[website] = false;
            }
        }

        private static readonly object _httpClientLock = new object();
        private static HttpClient _client;

        public static HttpClient Client
        {
            get
            {
                lock (_httpClientLock)
                {
                    if (_client == null)
                    {
                        _client = new HttpClient
                        {
                            Timeout = TimeSpan.FromSeconds(30)
                        };
                    }
                    return _client;
                }
            }
        }
        private static readonly object _accessChoiceLock = new object();
        private static RememberAccessChoice _accessChoice = new RememberAccessChoice();

        public static void ConfigureTLS()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls13;
        }

        public static void DisplayError(Exception ex)
        {
            MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void DisplayError(string messageIntro, Exception ex)
        {
            MessageBox.Show($"{messageIntro}: {ex.Message}", ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void DisplayWarning(Exception ex)
        {
            MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void DisplayInformation(Exception ex)
        {
            MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static bool GetInternetAccessPermission(Uri uri)
        {
            string baseUrl = uri.GetLeftPart(UriPartial.Authority);
            if (_accessChoice.IsGrantedAccess(baseUrl))
            {
                return true;
            }
            else
            {
                var result = MessageBox.Show(
                    $"{Forge.CultureHelper.GetLocalizedString("(CommonUtils.cs) [GetInternetAccessPermission] MessageBox #1 Text")}{Environment.NewLine}{baseUrl}",
                    Forge.CultureHelper.GetLocalizedString("(CommonUtils.cs) [GetInternetAccessPermission] MessageBox #1 Caption"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning
                );
                if (result == DialogResult.Yes)
                {
                    _accessChoice.Grant(baseUrl);
                    return true;
                }
                return false;
            }
        }

        public static Word.Application GetApplication()
        {
            return Globals.ThisAddIn.Application;
        }

        public static Document GetActiveDocument()
        {
            return Globals.ThisAddIn.Application.ActiveDocument;
        }

        public static Comments GetComments()
        {
            return GetActiveDocument().Comments;
        }

        public static Range GetSelectionRange()
        {
            return GetApplication().Selection.Range;
        }
        public static int GetWordPageCount()
        {
            int pageCount = GetActiveDocument().ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
            return pageCount;
        }

        public static IEnumerable<string> SplitString(string str, int chunkSize)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < str.Length; i += chunkSize)
            {
                if (i + chunkSize > str.Length)
                    chunkSize = str.Length - i;

                result.Add(str.Substring(i, chunkSize));
            }
            return result;
        }

        public static string SubstringTokens(string text, int maxTokens)
        {
            return SubstringWithoutBounds(text, TokensToCharCount(maxTokens));
        }
        private static string SubstringWithoutBounds(string text, int maxLen)
        {
            return (maxLen >= text.Length) ? text : text.Substring(0, maxLen);
        }

        public static int TokensToCharCount(int tokenCount)
        {
            return tokenCount * 4; // https://platform.openai.com/tokenizer
        }
        public static int CharToTokenCount(int charCount)
        {
            return charCount / 4; // https://platform.openai.com/tokenizer
        }

        public static void GetEnvironmentVariableIfAvailable(ref string dest, string variable)
        {
            var key = Environment.GetEnvironmentVariable(variable);
            if (key != null)
                dest = key;
        }
    }

    public class CultureLocalizationHelper
    {
        private ResourceManager _resourceManager;

        public CultureLocalizationHelper(string baseName, System.Reflection.Assembly assembly)
        {
            _resourceManager = new ResourceManager(baseName, assembly);
        }

        public string GetLocalizedString(string key, CultureInfo culture = null)
        {
            string localizedString = _resourceManager.GetString(key, culture);
            return (localizedString == null) ? _resourceManager.GetString(key, CultureInfo.InvariantCulture) : localizedString;
        }
    }
}
