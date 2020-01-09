using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using GenCode128;

namespace TournamentFormGenerator.ViewModel
{

    public class AnswersCollection
    {

        public static AnswersCollection BuildGroupedByTeam(int teamId, int RoundsCount, int QuestionsPerRound, string headerTemplate)
        {
            AnswersCollection result = new AnswersCollection();
            
            result.Headers = new Queue<QuestionForm>(RoundsCount * QuestionsPerRound);
            for (int roundId = 1; roundId <= RoundsCount; roundId++)
            {
                for (int queryId = 1; queryId <= QuestionsPerRound; queryId++)
                {
                    result.Headers.Enqueue(new QuestionForm()
                    {
                        QuestionId = queryId,
                        RoundId = roundId,
                        TeamId = teamId,
                        EndToEndQuestionId = (roundId - 1) * QuestionsPerRound + queryId,
                        HeaderTemplate = headerTemplate
                    });

                }
            }
            return result;
        }
        public static List<AnswersCollection> BuildGroupedByQuestion(int teamStart, int teamEnd, int RoundsCount, int QuestionsPerRound, string headerTemplate)
        {
            var result = new List<AnswersCollection>();

            for (int roundId = 1; roundId <= RoundsCount; roundId++)
            {
                for (int queryId = 1; queryId <= QuestionsPerRound; queryId++)
                {
                    var resultItem = new AnswersCollection();
                    resultItem.Headers = new Queue<QuestionForm>(RoundsCount * QuestionsPerRound);
                    for (int teamId = teamStart; teamId <= teamEnd; teamId++)
                    {
                        resultItem.Headers.Enqueue(new QuestionForm()
                        {
                            QuestionId = queryId,
                            RoundId = roundId,
                            TeamId = teamId,
                            EndToEndQuestionId = (roundId - 1) * QuestionsPerRound + queryId,
                            HeaderTemplate = headerTemplate
                        });
                    }
                    result.Add(resultItem);
                }
            }
            return result;
        }
        public Queue<QuestionForm> Headers { get; protected set; }
        
        public bool IsProcessed
        {
            get
            {
                return Headers.Count == 0;
            }
        }



        public void Print(Graphics g, int cellsOnXAxis, int cellsOnYAxis)
        {
            int smallFontSize = 36 / cellsOnXAxis;
            int largeFontSize = 200 / cellsOnXAxis;

            using (Font fnt = new Font("Arial", smallFontSize))
            {
                using (Font largeFnt = new Font("Arial", largeFontSize))
                {
                    using (var pen = new Pen(Color.Black, 1))
                    {
                        int cellWidth = 732 / cellsOnXAxis;
                        int celHeight = 1060 / cellsOnYAxis;

                        int startPosition = 35;
                        int margin = cellWidth * 57 / 1000; // ~14 (3x5)
                        int marginTop = celHeight * 132 / 1000; // ~ 28 (3x5)
                        int barMargin = 10;

                        for (int j = 0; j < cellsOnYAxis; j++)
                        {
                            for (int i = 0; i < cellsOnXAxis; i++)
                            {
                                if (!IsProcessed)
                                {
                                    var cellLeftPosition = i * cellWidth + startPosition;
                                    var cellTopPosition = j * celHeight + startPosition;

                                    var innerPosX = cellLeftPosition + margin;
                                    var innerWidth = cellWidth - 2 * margin;

                                    var captionPosX = innerPosX;
                                    var captionPosY = cellTopPosition + marginTop;

                                    var barPosX = innerPosX;
                                    var barPosY = captionPosY + smallFontSize + barMargin;
                                    var barHeight = celHeight / 8;
                                    var barWidth = innerWidth;

                                    // Outer rect
                                    g.DrawRectangle(pen, cellLeftPosition, cellTopPosition, cellWidth, celHeight);

                                    // Caption
                                    var questionFrm = Headers.Dequeue();
                                    g.DrawString(questionFrm.Caption, fnt, System.Drawing.Brushes.Black, captionPosX, captionPosY);

                                    // Draw in rectangle
                                    g.DrawImage(questionFrm.BarCode, new Rectangle(barPosX, barPosY, barWidth, barHeight));
                                    //g.DrawRectangle(pen, barPosX, barPosY, barWidth, barHeight);

                                    // Big number
                                    var numPosX = questionFrm.EndToEndQuestionId < 10 ? cellLeftPosition + cellWidth / 3 : cellLeftPosition + cellWidth / 4;
                                    var numPosY = cellTopPosition + celHeight / 3;
                                    g.DrawString(questionFrm.EndToEndQuestionId.ToString(), largeFnt, Brushes.LightGray, (int)numPosX, numPosY);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
