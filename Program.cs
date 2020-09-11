/*
 * SPDX-FileCopyrightText: 2020 Daniel Eder
 *
 * SPDX-License-Identifier: MIT
 */
using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;

namespace MultipleChoiceMaker
{
    class Program
    {
        public static Random Random = new Random();

        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: multiplechoicemaker.exe <input-excel> <output-excel>");
                return;
            }

            var inputFile = args[0];
            var outputFile = args[1];

            var choiceCount = 4;

            using (var inputWorkbook = new XLWorkbook(inputFile))
            using (var outputWorkbook = new XLWorkbook())
            {
                var inputWorksheet = inputWorkbook.Worksheet(1); //indices are 1-based in ClosedXML
                var outputWorksheet = outputWorkbook.AddWorksheet();

                //input worksheet should be in the format: 
                // | question | answer | 
                //output worksheet will be in the format:
                // | question     |  answer |
                // | - choice 1   |
                // | - choice 2   |
                // | - choice 3   |
                // | - choice 4   |
                // in HTML

                //Copy header
                outputWorksheet.Cell(1, 1).SetValue(inputWorksheet.Cell(1, 1).GetValue<String>());
                outputWorksheet.Cell(1, 2).SetValue(inputWorksheet.Cell(1, 2).GetValue<String>());

                var inputQuestions = inputWorksheet.Range(inputWorksheet.Cell(2, 1), inputWorksheet.Column(1).LastCellUsed()); //skip header
                var inputAnswers = inputWorksheet.Range(inputWorksheet.Cell(2, 2), inputWorksheet.Column(2).LastCellUsed()); //skip header

                var outputQuestions = outputWorksheet.Range(outputWorksheet.Cell(2, 1), outputWorksheet.Column(1).LastCell()); //skip header
                var outputAnswers = outputWorksheet.Range(outputWorksheet.Cell(2, 2), outputWorksheet.Column(2).LastCell()); //skip header

                foreach (var inputQuestionCell in inputQuestions.Cells())
                {
                    var row = inputQuestionCell.Address.RowNumber-1; //rownumber is absolute address, but we are within the relative range, so we have to subtract the header row
                    var inputQuestion = inputQuestionCell.GetValue<string>();
                    var inputAnswer = inputAnswers.Cell(row, 1).GetValue<string>();
                    outputAnswers.Cell(row, 1).SetValue(inputAnswer); //copy answer 1:1

                    var choices = SelectRandomAnswers(inputAnswers, row, choiceCount-1);
                    choices.Add(inputAnswer);

                    Shuffle(choices);

                    var questionBuilder = new StringBuilder();
                    questionBuilder.Append($"<p><b>{inputQuestion}:</b></p>");
                    questionBuilder.Append("<ul>");
                    foreach(var choice in choices)
                    {
                        questionBuilder.Append($"<li>{choice}</li>");
                    }
                    questionBuilder.Append("</ul>");
                    outputQuestions.Cell(row, 1).SetValue(questionBuilder.ToString());
                }

                outputWorkbook.SaveAs(outputFile);
            }
        }

        /// <summary>
        /// Selects random answers from the given answer range, skipping the excludedIndex.
        /// </summary>
        /// <param name="answers">The range to select an answer from.</param>
        /// <param name="excludedIndex">The cell index to skip (the correct answer that should not be included in the random ones).</param>
        /// <param name="count">The count of answers to select</param>
        /// <returns>A list of answer strings.</returns>
        public static List<String> SelectRandomAnswers(IXLRange answers, int excludedIndex, int count)
        {
            var output = new List<string>();
            var exclusions = new List<int>() { excludedIndex };

            //Get range for random
            int min = answers.FirstCellUsed().Address.RowNumber;
            int max = answers.LastCellUsed().Address.RowNumber;

            //iterate until we have the desired amount of results
            while (output.Count < count)
            {
                int selectedIndex = Random.Next(min, max);

                //Skip the correct answer, and avoid duplicate random answers
                if (exclusions.Contains(selectedIndex))
                    continue;

                output.Add(answers.Cell(selectedIndex, 1).GetValue<string>());
                exclusions.Add(selectedIndex);
            }

            return output;
        }

        /// <summary>
        /// Fisher-Yates shuffle
        /// </summary>
        static void Shuffle<T>(List<T> list)
        {
            int n = list.Count;
            for (int i = 0; i < n; i++)
            {
                int r = i + (int)(Random.NextDouble() * (n - i));
                T t = list[r];
                list[r] = list[i];
                list[i] = t;
            }
        }
    }
}
