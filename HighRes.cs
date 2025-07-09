using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

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
                    ShowSummary(skipReasons, "No structures need conversion:");
                    return;
                }

                // Show conversion confirmation
                string conversionList = string.Join("\n", toConvert.Select(s => s.Id));
                string message = $"{toConvert.Count} structures will be converted to high resolution:\n\n{conversionList}\n\n" +
                                "This operation cannot be undone. Proceed?";

                if (MessageBox.Show(message, "Confirm Conversion", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                    return;

                // Perform conversion
                context.Patient.BeginModifications();
                ConvertStructures(toConvert, skipReasons);
                ShowSummary(skipReasons, "Conversion complete:");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Critical error: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void ConvertStructures(List<Structure> toConvert, List<string> skipReasons)
        {
            int converted = 0;
            foreach (Structure s in toConvert)
            {
                try
                {
                    s.ConvertToHighResolution();
                    converted++;
                    skipReasons.Add($"{s.Id}: Successfully converted");
                }
                catch (Exception ex)
                {
                    skipReasons.Add($"{s.Id}: FAILED - {ex.Message}");
                }
            }
            skipReasons.Insert(0, $"Successfully converted {converted}/{toConvert.Count} structures");
        }

        private void ShowSummary(List<string> messages, string title)
        {
            string report = string.Join("\n", messages);
            MessageBox.Show(report, title);
        }
    }
}