using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;

namespace TextForge.Services
{
    public class ErrorHandlingService
    {
        private readonly ILogger _logger;

        public ErrorHandlingService(ILogger<ErrorHandlingService> logger)
        {
            _logger = logger;
        }

        public void HandleError(Exception ex, string context = null)
        {
            _logger.LogError(ex, "Error occurred: {Message}", ex.Message);
            
            var errorMessage = context != null 
                ? $"{context}: {ex.Message}"
                : ex.Message;

            ShowErrorMessage(errorMessage, ex.GetType().Name);
        }

        public void HandleWarning(Exception ex, string context = null)
        {
            _logger.LogWarning(ex, "Warning: {Message}", ex.Message);
            
            var warningMessage = context != null
                ? $"{context}: {ex.Message}"
                : ex.Message;

            ShowWarningMessage(warningMessage, ex.GetType().Name);
        }

        public void HandleInformation(string message, string title = "Information")
        {
            _logger.LogInformation(message);
            ShowInformationMessage(message, title);
        }

        private void ShowErrorMessage(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowWarningMessage(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowInformationMessage(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public bool ShouldRetry(Exception ex, int retryCount)
        {
            // Add logic to determine if operation should be retried
            return retryCount < 3 && !(ex is OperationCanceledException);
        }

        public void LogOperation(string operationName, TimeSpan duration)
        {
            _logger.LogInformation("Operation {OperationName} completed in {Duration}ms", 
                operationName, duration.TotalMilliseconds);
        }
    }
}
