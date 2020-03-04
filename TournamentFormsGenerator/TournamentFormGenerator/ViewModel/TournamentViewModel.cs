using GalaSoft.MvvmLight;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using GenCode128;
using System.Threading;
using System.IO;
using Microsoft.Win32;
using System;
using System.Drawing.Printing;
namespace TournamentFormGenerator.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// Use the <strong>mvvminpc</strong> snippet to add bindable properties to this ViewModel.
    /// </para>
    /// <para>
    /// You can also use Blend to data bind with the tool's support.
    /// </para>
    /// <para>
    /// See http://www.galasoft.ch/mvvm/getstarted
    /// </para>
    /// </summary>
    public class TournamentViewModel : ViewModelBase
    {
        #region properties


        /// <summary>
        /// The <see cref="RoundsCount" /> property's name.
        /// </summary>
        public const string RoundsCountPropertyName = "RoundsCount";

        private int _RoundsCount = 3;

        /// <summary>
        /// Gets the RoundsCount property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public int RoundsCount
        {
            get
            {
                return _RoundsCount;
            }
            set
            {
                if (_RoundsCount == value)
                    return;
                _RoundsCount = value;
                RaisePropertyChanged(RoundsCountPropertyName);
            }
        }
        /// <summary>
        /// The <see cref="TeamsRangeBegin" /> property's name.
        /// </summary>
        public const string TeamsRangeBeginPropertyName = "TeamsRangeBegin";

        private int _TeamsRangeBegin = 1;

        /// <summary>
        /// Gets the TeamsRangeBegin property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public int TeamsRangeBegin
        {
            get
            {
                return _TeamsRangeBegin;
            }
            set
            {
                if (_TeamsRangeBegin == value)
                    return;
                _TeamsRangeBegin = value;

                if (TeamsRangeBegin > TeamsRangeEnd)
                    TeamsRangeEnd = TeamsRangeBegin;

                RaisePropertyChanged(TeamsRangeBeginPropertyName);
            }
        }

        /// <summary>
        /// The <see cref="TeamsRangeEnd" /> property's name.
        /// </summary>
        public const string TeamsRangeEndPropertyName = "TeamsRangeEnd";

        private int _TeamsRangeEnd = 12;

        /// <summary>
        /// Gets the TeamsRangeEnd property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public int TeamsRangeEnd
        {
            get
            {
                return _TeamsRangeEnd;
            }
            set
            {
                if (_TeamsRangeEnd == value)
                    return;
                _TeamsRangeEnd = value;

                if (TeamsRangeBegin > TeamsRangeEnd)
                    TeamsRangeBegin = TeamsRangeEnd;
                RaisePropertyChanged(TeamsRangeEndPropertyName);
            }
        }

        List<AnswersCollection> QuestionForms;
        /// <summary>
        /// The <see cref="QuestionsPerRound" /> property's name.
        /// </summary>
        public const string QuestionsPerRoundPropertyName = "QuestionsPerRound";

        private int _QuestionsPerRound = 18;

        /// <summary>
        /// Gets the QuestionsPerRound property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public int QuestionsPerRound
        {
            get
            {
                return _QuestionsPerRound;
            }
            set
            {
                if (_QuestionsPerRound == value)
                    return;
                _QuestionsPerRound = value;
                RaisePropertyChanged(QuestionsPerRoundPropertyName);
            }
        }

        /// <summary>
        /// The <see cref="IsProcessing" /> property's name.
        /// </summary>
        public const string IsProcessingPropertyName = "IsProcessing";

        private bool _IsProcessing = false;

        /// <summary>
        /// Gets the IsProcessing property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public bool IsProcessing
        {
            get
            {
                return _IsProcessing;
            }
            set
            {
                if (_IsProcessing == value)
                    return;
                _IsProcessing = value;
                RaisePropertyChanged(IsProcessingPropertyName);
                PrintCommand.RaiseCanExecuteChanged();
            }
        }

        /// <summary>
        /// The <see cref="PrintingProgress" /> property's name.
        /// </summary>
        public const string PrintingProgressPropertyName = "PrintingProgress";

        private int _PrintingProgress = 0;

        /// <summary>
        /// Gets the PrintingProgress property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public int PrintingProgress
        {
            get
            {
                return _PrintingProgress;
            }
            set
            {
                if (_PrintingProgress == value)
                    return;
                _PrintingProgress = value;
                RaisePropertyChanged(PrintingProgressPropertyName);
            }
        }

        #endregion
        public GalaSoft.MvvmLight.Command.RelayCommand PrintCommand { get; protected set; }
        public GalaSoft.MvvmLight.Command.RelayCommand OpenSpreadsheetCommand { get; protected set; }
        /// <summary>
        /// Initializes a new instance of the TournamentViewModel class.
        /// </summary>
        public TournamentViewModel()
        {
            PrinterSettings.InstalledPrinters.Cast<string>();
            _RoundsCount = 3;
            _QuestionsPerRound = 12;
            OpenSpreadsheetCommand = new GalaSoft.MvvmLight.Command.RelayCommand(() =>
                {
                    try
                    {
                        var dlg = new SaveFileDialog();
                        dlg.FileName = "tournament.ods";
                        dlg.AddExtension = true;
                        dlg.DefaultExt = "*.ods";
                        dlg.CheckFileExists = false;
                        dlg.Filter = "OpenOffice Calc Spreadsheet|*.ods";
                        if (dlg.ShowDialog().Value)
                        {
                            var fileData = Properties.Resources.tournamentFinal;
                            var path = dlg.FileName;
                            File.WriteAllBytes(path, fileData);
                        }

                    }
                    catch (Exception ex)
                    {
                        ShowError("Не удалось сохранить файл. " + ex.ToString());
                    }
                });
            PrintCommand = new GalaSoft.MvvmLight.Command.RelayCommand(() =>
                {
                    IsProcessing = true;
                    ThreadPool.QueueUserWorkItem(new WaitCallback((o) =>
                        {
                            try
                            {
                                QuestionForms = new List<AnswersCollection>();

                                if (GroupByQuestion)
                                {
                                    QuestionForms =
                                        AnswersCollection.BuildGroupedByQuestion(TeamsRangeBegin, TeamsRangeEnd, RoundsCount, QuestionsPerRound, ItemHeaderTemplate);
                                }
                                else
                                {
                                    for (int teamId = TeamsRangeBegin; teamId <= TeamsRangeEnd; teamId++)
                                        QuestionForms.Add(
                                            AnswersCollection.BuildGroupedByTeam(teamId, RoundsCount, QuestionsPerRound, ItemHeaderTemplate));
                                }


                                using (var printDocument1 = new System.Drawing.Printing.PrintDocument())
                                {
                                    printDocument1.PrinterSettings.PrinterName = TargetPrinter;
                                    printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printDocument1_PrintPage);
                                    printDocument1.Print();
                                }
                                IsProcessing = false;

                            }
                            catch (Exception ex)
                            {
                                ShowError("Не удалось напечатать документ. " + ex.ToString());
                            }
                        }));

                }, () => !IsProcessing);
            if (IsInDesignMode)
            {
                // Code runs in Blend --> create design time data.
            }
            else
            {
                // Code runs "for real"
            }
        }

        private void ShowError(string msg)
        {
            this.IsProcessing = true;
            Message = msg;
            ThreadPool.QueueUserWorkItem(new WaitCallback((o) =>
                {

                    Thread.Sleep(5000);
                    IsProcessing = false;
                    Message = DefaultWaitingMessage;
                }
                ));
        }

        void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            using (Graphics g = e.Graphics)
            {

                var form = QuestionForms.FirstOrDefault(t => !t.IsProcessed);
                if (form != null)
                {
                    var width = UseSize_4x9 ? 4 : 3;
                    var height = UseSize_4x9 ? 9 : 6;
                    form.Print(g, width, height);
                }
                PrintingProgress = QuestionForms.Count - QuestionForms.Count(t => !t.IsProcessed);

                e.HasMorePages = QuestionForms.Any(t => !t.IsProcessed);

            }
        }
        const string DefaultWaitingMessage = "Подождите, идет обработка";

        /// <summary>
        /// The <see cref="Message" /> property's name.
        /// </summary>
        public const string MessagePropertyName = "Message";

        private string _Message = DefaultWaitingMessage;

        /// <summary>
        /// Gets the Message property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public string Message
        {
            get
            {
                return _Message;
            }
            set
            {
                if (_Message == value)
                    return;
                _Message = value;
                RaisePropertyChanged(MessagePropertyName);
            }
        }


        /// <summary>
        /// The <see cref="TargetPrinter" /> property's name.
        /// </summary>
        public const string TargetPrinterPropertyName = "TargetPrinter";

        private string _TargetPrinter = "Microsoft XPS Document Writer";

        /// <summary>
        /// Gets the TargetPrinter property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public string TargetPrinter
        {
            get
            {
                return _TargetPrinter;
            }
            set
            {
                if (_TargetPrinter == value)
                    return;
                _TargetPrinter = value;
                RaisePropertyChanged(TargetPrinterPropertyName);
            }
        }


        /// <summary>
        /// The <see cref="UseSize_4x9" /> property's name.
        /// </summary>
        public const string UseSize_4x9PropertyName = "UseSize_4x9";

        private bool _UseSize_4x9 = false;

        /// <summary>
        /// Gets the UseSize_4x9 property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public bool UseSize_4x9
        {
            get
            {
                return _UseSize_4x9;
            }
            set
            {
                if (_UseSize_4x9 == value)
                    return;
                _UseSize_4x9 = value;
                RaisePropertyChanged(UseSize_4x9PropertyName);
            }
        }



        public const string GroupByQuestionPropertyName = "GroupByQuestion";

        private bool _GroupByQuestion = false;

        /// <summary>
        /// Gets the GroupByQuestion property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public bool GroupByQuestion
        {
            get
            {
                return _GroupByQuestion;
            }
            set
            {
                if (_GroupByQuestion == value)
                    return;
                _GroupByQuestion = value;
                RaisePropertyChanged(GroupByQuestionPropertyName);
            }
        }





        public const string ItemHeaderTemplatePropertyName = "ItemHeaderTemplate";

        private string _ItemHeaderTemplate = "Тур {0} Вопрос {1} Команда {2}";

        /// <summary>
        /// Gets the ItemHeaderTemplate property.
        /// Property's value raise the PropertyChanged event.
        /// </summary>
        public string ItemHeaderTemplate
        {
            get
            {
                return _ItemHeaderTemplate;
            }
            set
            {
                if (_ItemHeaderTemplate == value)
                    return;
                _ItemHeaderTemplate = value;
                RaisePropertyChanged(ItemHeaderTemplatePropertyName);
            }
        }

    }
}