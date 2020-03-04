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
            using (Font fnt = new Font("Arial", 36 / cellsOnXAxis))
            {
                using (Font largeFnt = new Font("Arial", 200 / cellsOnXAxis))
                {
                    using (var pen = new Pen(Color.Black, 1))
                    {
                        int startPosition = 35;
                        int margin = 4;
                        int cellWidth = 732 / cellsOnXAxis;
                        int celHeight = 1060 / cellsOnYAxis;
                        for (int j = 0; j < cellsOnYAxis; j++)
                        {
                            for (int i = 0; i < cellsOnXAxis; i++)
                            {

                                if (!IsProcessed)
                                {
                                    var cellLeftPosition = i * cellWidth + startPosition;
                                    var cellTopPosition = j * celHeight + startPosition;
                                    g.DrawRectangle(pen, cellLeftPosition, cellTopPosition, cellWidth, celHeight);
                                    var questionFrm = Headers.Dequeue();
                                    g.DrawString(questionFrm.Caption, fnt, System.Drawing.Brushes.Black, cellLeftPosition + margin, cellTopPosition + margin);
                                    //g.DrawImage(questionFrm.BarCode, cellLeftPosition + cellWidth / 7, cellTopPosition + celHeight / 5);
                                    // Draw in rectangle
                                    var barX = cellLeftPosition + cellWidth / 18;
                                    var barY = cellTopPosition + celHeight / 6;
                                    g.DrawImage(questionFrm.BarCode, new Rectangle(barX, barY, cellWidth * 3 / 4, celHeight / 8));


                                    var xPos = questionFrm.EndToEndQuestionId < 10 ? cellLeftPosition + cellWidth / 3 : cellLeftPosition + cellWidth / 4;
                                    g.DrawString(questionFrm.EndToEndQuestionId.ToString(), largeFnt, Brushes.LightGray, (int)xPos, cellTopPosition + celHeight / 3);
                                }
                            }
                        }
                    }
                }



            }
        }
    }
}
