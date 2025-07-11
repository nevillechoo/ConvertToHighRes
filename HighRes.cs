using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyVersion("1.0.0.1")]
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public void Execute(ScriptContext context)
        {
            try
            {
                // Get current structure set
                StructureSet structureSet = context.StructureSet;
                if (structureSet == null)
                {
                    MessageBox.Show("No structure set loaded. Please open a structure set.");
                    return;
                }

                // Identify convertible structures
                List<Structure> toConvert = new List<Structure>();
                List<string> skipReasons = new List<string>();

                foreach (Structure s in structureSet.Structures)
                {
                    if (s.IsHighResolution)
                    {
                        skipReasons.Add($"{s.Id}: Already high resolution");
                    }
                    else if (s.IsApproved)
                    {
                        skipReasons.Add($"{s.Id}: Unable to convert approved structures");
                    }
                    else if (!s.CanConvertToHighResolution())
                    {
                        skipReasons.Add($"{s.Id}: Unable to convert to high resolution");
                    }
                    else if (s.Id.Contains("ETD"))
                    {
                        skipReasons.Add($"{s.Id}: ETD structures do not require conversion");
                    }
                    else if (s.Id.Contains("zOrig"))
                    {
                        skipReasons.Add($"{s.Id}: zOrig structures do not require conversion");
                    }
                    //else if (s.IsEmpty)
                    //{
                    //    skipReasons.Add($"{s.Id}: Empty structure (no contours)");
                    //}
                    else if (s.DicomType == "SUPPORT" || s.DicomType == "MARKER")
                    {
                        skipReasons.Add($"{s.Id}: Unsupported DICOM type ({s.DicomType})");
                    }
                    else
                    {
                        toConvert.Add(s);
                    }
                }

                // Check if any structures need conversion
                if (toConvert.Count == 0)
                {
                    ShowScrollableMessage(skipReasons, "No structures need conversion:");
                    return;
                }

                // Show conversion confirmation
                string conversionList = string.Join("\n", toConvert.Select(s => s.Id));
                string message = $"{toConvert.Count} structures will be converted to high resolution:\n\n{conversionList}\n\n" +
                                "This operation cannot be undone. Proceed?";

                if (ShowScrollableConfirmation(message, "Confirm Conversion") != MessageBoxResult.Yes)
                    return;

                // Perform conversion
                context.Patient.BeginModifications();
                ConvertStructures(toConvert, skipReasons);
                ShowScrollableMessage(skipReasons, "Conversion complete:");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Critical error: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void ConvertStructures(List<Structure> toConvert, List<string> skipReasons)
        {
            int converted = 0;

            // Create progress window
            var progressWindow = CreateProgressWindow(toConvert.Count);
            progressWindow.Show();

            foreach (Structure s in toConvert)
            {
                try
                {
                    s.ConvertToHighResolution();
                    converted++;
                    skipReasons.Add($"{s.Id}: Successfully converted");

                    // Update progress
                    UpdateProgress(progressWindow, converted, toConvert.Count);
                }
                catch (Exception ex)
                {
                    skipReasons.Add($"{s.Id}: FAILED - {ex.Message}");
                }
            }

            // Close progress window after a brief delay
            System.Threading.Thread.Sleep(500);
            progressWindow.Close();

            skipReasons.Insert(0, $"Successfully converted {converted}/{toConvert.Count} structures");
        }

        private Window CreateProgressWindow(int totalStructures)
        {
            var window = new Window
            {
                Title = "Converting Structures",
                Width = 300,
                Height = 100,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Topmost = true
            };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition());
            grid.RowDefinitions.Add(new RowDefinition());

            var label = new Label
            {
                Content = "Converting structures to high resolution...",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            var progressLabel = new Label
            {
                Name = "ProgressLabel",
                Content = $"0/{totalStructures}",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14,
                FontWeight = FontWeights.Bold
            };

            Grid.SetRow(label, 0);
            Grid.SetRow(progressLabel, 1);

            grid.Children.Add(label);
            grid.Children.Add(progressLabel);

            window.Content = grid;
            return window;
        }

        private void UpdateProgress(Window progressWindow, int current, int total)
        {
            var grid = progressWindow.Content as Grid;
            var progressLabel = grid.Children.OfType<Label>().FirstOrDefault(l => l.Name == "ProgressLabel");

            if (progressLabel != null)
            {
                progressLabel.Content = $"{current}/{total}";
            }

            // Force UI update
            progressWindow.Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
        }

        private void ShowScrollableMessage(List<string> messages, string title)
        {
            string report = string.Join("\n", messages);
            ShowScrollableDialog(report, title, false);
        }

        private MessageBoxResult ShowScrollableConfirmation(string message, string title)
        {
            return ShowScrollableDialog(message, title, true);
        }

        private MessageBoxResult ShowScrollableDialog(string content, string title, bool isConfirmation)
        {
            var window = new Window
            {
                Title = title,
                Width = 500,
                Height = 400,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.CanResize
            };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Scrollable text area
            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(10)
            };

            var textBlock = new TextBlock
            {
                Text = content,
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new System.Windows.Media.FontFamily("Verdana"),
                FontSize = 11
            };

            scrollViewer.Content = textBlock;
            Grid.SetRow(scrollViewer, 0);
            grid.Children.Add(scrollViewer);

            // Button panel
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(10)
            };

            MessageBoxResult result = MessageBoxResult.OK;

            if (isConfirmation)
            {
                var yesButton = new Button
                {
                    Content = "Yes",
                    Width = 75,
                    Height = 25,
                    Margin = new Thickness(5, 0, 0, 0)
                };
                yesButton.Click += (s, e) => { result = MessageBoxResult.Yes; window.Close(); };

                var noButton = new Button
                {
                    Content = "No",
                    Width = 75,
                    Height = 25,
                    Margin = new Thickness(5, 0, 0, 0)
                };
                noButton.Click += (s, e) => { result = MessageBoxResult.No; window.Close(); };

                buttonPanel.Children.Add(yesButton);
                buttonPanel.Children.Add(noButton);
            }
            else
            {
                var okButton = new Button
                {
                    Content = "OK",
                    Width = 75,
                    Height = 25
                };
                okButton.Click += (s, e) => { result = MessageBoxResult.OK; window.Close(); };

                buttonPanel.Children.Add(okButton);
            }

            Grid.SetRow(buttonPanel, 1);
            grid.Children.Add(buttonPanel);

            window.Content = grid;
            window.ShowDialog();

            return result;
        }
    }
}