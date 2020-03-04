using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using GenCode128;
namespace TournamentFormGenerator.ViewModel
{
    public class QuestionForm
    {
        public string Caption
        {
            get
            {
                return string.Format(HeaderTemplate, RoundId, QuestionId, TeamId);
            }
        }
        public Image BarCode
        {
            get
            {
                //var result = Code128Rendering.MakeBarcodeImage(RoundId.ToString() + QuestionId.ToString("D2") + TeamId.ToString("D2"), 1, true);
                // Make it wider and shring on drawing
                var result = Code128Rendering.MakeBarcodeImage(RoundId.ToString() + QuestionId.ToString("D2") + TeamId.ToString("D2"), 2, true);
                return result;
            }
        }

        public int RoundId { get; set; }
        public int QuestionId { get; set; }
        public int TeamId { get; set; }

        public int EndToEndQuestionId { get; set; }

        public string HeaderTemplate { get; set; }
    }
}
