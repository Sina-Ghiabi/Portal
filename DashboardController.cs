using Microsoft.AspNetCore.Mvc;
using Abp.AspNetCore.Mvc.Authorization;
using Dapna.MSVPortal.Controllers;
using Dapna.MSVPortal.Projects;
using Dapna.MSVPortal.Financial;
using System.Threading.Tasks;
using Dapna.MSVPortal.Web.ViewModels;
using System.Linq;
using Abp.Web.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Dapna.MSVPortal.Enums;
using MD.PersianDateTime;
using System;
using System.Collections.Generic;
using Dapna.MSVPortal.Web.ViewModels.ChartForm;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.AspNetCore.Mvc.Rendering;
using Abp.Extensions;
using System.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Dapna.MSVPortal.ProjectsDocumentations;

namespace Dapna.MSVPortal.Web.Controllers
{
    [AbpMvcAuthorize]
    public class DashboardController : MSVPortalControllerBase
    {
        private readonly IPaymentRequestAppService _paymentRequestAppService;
        private readonly IReceiveBillAppService _receiveBillAppService;
        private readonly ITransactionAppService _transactionAppService;
        private readonly IProjectAppService _projectAppService;
        private readonly IContractorAppService _contractorAppService;
        private readonly IProject_User_MappingAppService _project_User_MappingAppService;
        private readonly ICostBenefitAppService _costBenefitAppService;
        private readonly ILaborInventoryAppService _laborInventoryAppService;
        private readonly IReceiveAndPaymentAppService _receiveAndPaymentAppService;
        private readonly IWarrantyBalanceAppService _warrantyBalanceAppService;
        private readonly IProgressPercentageAppService _progressPercentageAppService;
        private readonly IGrossProfitToNetAppService _grossProfitToNetAppService;
        private readonly IProjectsTasksAppService _projectsTasksAppService;
        private readonly IProjectsTasksRevisionsAppService _projectsTasksRevisionsAppService;

        public decimal itemTotal { get; private set; }

        public DashboardController(IPaymentRequestAppService paymentRequestAppService,
            IReceiveBillAppService receiveBillAppService,
            ITransactionAppService transactionAppService,
            IProjectAppService projectAppService,
            IContractorAppService contractorAppService,
            IProject_User_MappingAppService project_User_MappingAppService,
            ICostBenefitAppService costBenefitAppService,
            ILaborInventoryAppService laborInventoryAppService,
            IReceiveAndPaymentAppService receiveAndPaymentAppService,
            IWarrantyBalanceAppService warrantyBalanceAppService,
            IProgressPercentageAppService progressPercentageAppService,
            IGrossProfitToNetAppService grossProfitToNetAppService,
            IProjectsTasksAppService projectsTasksAppService,
            IProjectsTasksRevisionsAppService projectsTasksRevisionsAppService
            )
        {
            _paymentRequestAppService = paymentRequestAppService;
            _receiveBillAppService = receiveBillAppService;
            _transactionAppService = transactionAppService;
            _projectAppService = projectAppService;
            _contractorAppService = contractorAppService;
            _project_User_MappingAppService = project_User_MappingAppService;
            _costBenefitAppService = costBenefitAppService;
            _laborInventoryAppService = laborInventoryAppService;
            _receiveAndPaymentAppService = receiveAndPaymentAppService;
            _warrantyBalanceAppService = warrantyBalanceAppService;
            _progressPercentageAppService = progressPercentageAppService;
            _grossProfitToNetAppService = grossProfitToNetAppService;
            _projectsTasksAppService = projectsTasksAppService;
            _projectsTasksRevisionsAppService = projectsTasksRevisionsAppService;
        }

        public async Task<ActionResult> Projects(int? id)
        {
            var userId = AbpSession.UserId.Value;

            var items = await _project_User_MappingAppService.GetUserProjects(userId);

            var records = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            ViewBag.ProjectId = new SelectList(records, "Id", "Title", id);

            return View();
        }


        public async Task<IActionResult> ProjectsDocumentations()
        {
            var Projects = await _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value);

            ViewBag.ProjectId = Projects.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            return View();
        }

        [DontWrapResult]
        public async Task<ActionResult> AFCDocuments(int? ProjectID , string Year = "")
        {
            DateTime? ConvertedStartYear = null;
            DateTime? ConvertedEndYear = null;

            if (Year != null)
            {
                string StartDate = Year + "/01/01";
                string EndDate = Year + "/12/29";

                ConvertedStartYear = PersianDateTime.Parse(StartDate);
                ConvertedEndYear = PersianDateTime.Parse(EndDate);
            }

            var Info = await _projectsTasksAppService.AFCDocumentsFilter(ProjectID , ConvertedStartYear , ConvertedEndYear);

            string[] DiciplineTypes = { "Civil", "Electrical", "Instrument", "Mechanical", "Piping", "Process", "Quality" };

            int[] DiciplineValues = {
                Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)1).Count(),
                Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)2).Count(),
                Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)3).Count(),
                Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)4).Count(),
                Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)5).Count(),
                Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)6).Count(),
                Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)7).Count()
            };

            return Json(new { DiciplineTypes , DiciplineValues }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [DontWrapResult]
        public async Task<ActionResult> AllDocuments(int? ProjectID, string Year = "", string Month = "")
        {
            string StartDate = "";
            string EndDate = "";

            DateTime? ConvertedStartYear = null;
            DateTime? ConvertedEndYear = null;

            var List = new List<int>();
            var Result = new List<ProjectsTasks>();

            if (Year != null && Month != null)
            {
                if (Month == "12" || Month == "11" || Month == "10") { StartDate = Year + "/" + Month + "/01"; } else { StartDate = Year + "/0" + Month + "/01"; };
                if (Month == "12") { EndDate = Year + "/" + Month + "/29"; } else if (Month == "11" || Month == "10") { EndDate = Year + "/" + Month + "/30"; } else if( Month == "9" || Month == "8" || Month == "7") { EndDate = Year + "/0" + Month + "/30"; } else { EndDate = Year + "/0" + Month + "/31"; }

                ConvertedStartYear = PersianDateTime.Parse(StartDate);
                ConvertedEndYear = PersianDateTime.Parse(EndDate);
            }
            else if (Year != null && Month == null)
            {
                StartDate = Year + "/01/01";
                EndDate = Year + "/12/29";

                ConvertedStartYear = PersianDateTime.Parse(StartDate);
                ConvertedEndYear = PersianDateTime.Parse(EndDate);
            }

            if (ProjectID.HasValue) 
            {
                var DocumentsSearch = await _projectsTasksAppService.GetAllProjectsTasks();
                var DocumentsFilter = DocumentsSearch.Where(a => a.ProjectID == ProjectID);
                var DocumentsResult = DocumentsFilter.Select(a => a.Id).ToList();

                foreach (var Item in DocumentsResult) { List.Add(Item); }
            }
            else
            {
               var DocumentSearch = await _projectsTasksAppService.GetAllProjectsTasks();
               var DocumentsResult = DocumentSearch.Select(a => a.Id).ToList();

               foreach (var Item in DocumentsResult) { List.Add(Item); }
            }

            var RevisionsFilter = await _projectsTasksRevisionsAppService.AllDocumentsFilter( List, ConvertedStartYear , ConvertedEndYear);

            var RevisionsResult = RevisionsFilter.Select(a => a.TaskID).ToList();

            foreach (var Item in RevisionsResult)
            {
                var DocumentResult = _projectsTasksAppService.GetSpecificProjectsTask(Item).Result.FirstOrDefault();

                Result.Add(new ProjectsTasks
                {
                    Id = DocumentResult.Id,
                    ProjectID = DocumentResult.ProjectID,
                    ProjectCode = DocumentResult.ProjectCode,
                    CompanyName = DocumentResult.CompanyName,
                    DocumentTitle = DocumentResult.DocumentTitle,
                    DocumentNumber = DocumentResult.DocumentNumber,
                    Dicipline = DocumentResult.Dicipline,
                    ResponsiblePerson = DocumentResult.ResponsiblePerson,
                    DocumentType = DocumentResult.DocumentType,
                    Description = DocumentResult.Description,
                    WeightFactor = DocumentResult.WeightFactor,
                    Progress = DocumentResult.Progress,
                    OriginalDuration = DocumentResult.OriginalDuration,
                    SourceOfItem = DocumentResult.SourceOfItem,
                    ManPower = DocumentResult.ManPower,
                    Critical = DocumentResult.Critical
                });
            }

            string[] DiciplineTypes = { "Civil", "Electrical", "Instrument", "Mechanical", "Piping", "Process", "Quality" };

            int[] DiciplineValues = {
                Result.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)1).Count(),
                Result.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)2).Count(),
                Result.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)3).Count(),
                Result.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)4).Count(),
                Result.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)5).Count(),
                Result.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)6).Count(),
                Result.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)7).Count()
            };

            return Json(new { DiciplineTypes, DiciplineValues , }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [DontWrapResult]
        public async Task<ActionResult> AverageAFCDocuments(int? ProjectID, string Year = "")
        {

            List<int> DiciplineValues = new List<int>();

            DateTime? ConvertedStartYear = null;
            DateTime? ConvertedEndYear = null;


            ProjectsTasksViewModel.AverageAFCDocuments AverageAFCDocuments = delegate (List<ProjectsTasks> List , int DiciplineType , int? ProjectIDs , DateTime? StartDate , DateTime? EndDate)
            {
                var RevisionsTotal = 0;

                var DocumentsIDs = List.Select(a => a.Id);

                foreach(var ID in DocumentsIDs)
                {
                    var Revision = _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(ID);

                    var RevisionNumber = Convert.ToInt32(Revision.Result.Select(a => a.RevisionNumber).LastOrDefault());

                    RevisionsTotal = RevisionsTotal + RevisionNumber;
                }

                var AllDocuments = _projectsTasksAppService.SelectedAFCDocument(ProjectIDs, StartDate, EndDate);

                var Result = (float)RevisionsTotal / (float)AllDocuments.Result.Where(a => a.Dicipline ==(ProjectsTasksDiciplineTypes)DiciplineType && a.LastStatus == (ProjectsTasksStatusTypes)1).Count();

                DiciplineValues.Add((int)Result);
            };

            if (Year != null )
            {
                var StartDate = Year + "/01/01";
                var EndDate = Year + "/12/29";

                ConvertedStartYear = PersianDateTime.Parse(StartDate);
                ConvertedEndYear = PersianDateTime.Parse(EndDate);
            }

            var Info = await _projectsTasksAppService.AverageAFCDocuments(ProjectID , ConvertedStartYear , ConvertedEndYear);

            string[] DiciplineTypes = { "Civil", "Electrical", "Instrument", "Mechanical", "Piping", "Process", "Quality" };

            if (Info.Count() > 0)
            {
                
                AverageAFCDocuments(Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)1 && a.LastStatus == (ProjectsTasksStatusTypes)1).ToList() , 1 , ProjectID, ConvertedStartYear , ConvertedEndYear); //Civil
                AverageAFCDocuments(Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)2 && a.LastStatus == (ProjectsTasksStatusTypes)1).ToList() , 2 , ProjectID, ConvertedStartYear , ConvertedEndYear); //Electrical
                AverageAFCDocuments(Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)3 && a.LastStatus == (ProjectsTasksStatusTypes)1).ToList() , 3 , ProjectID, ConvertedStartYear , ConvertedEndYear); //Instrument
                AverageAFCDocuments(Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)4 && a.LastStatus == (ProjectsTasksStatusTypes)1).ToList() , 4 , ProjectID, ConvertedStartYear , ConvertedEndYear); //Mechanical
                AverageAFCDocuments(Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)5 && a.LastStatus == (ProjectsTasksStatusTypes)1).ToList() , 5 , ProjectID, ConvertedStartYear , ConvertedEndYear); //Piping
                AverageAFCDocuments(Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)6 && a.LastStatus == (ProjectsTasksStatusTypes)1).ToList() , 6 , ProjectID, ConvertedStartYear , ConvertedEndYear); //Process
                AverageAFCDocuments(Info.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)7 && a.LastStatus == (ProjectsTasksStatusTypes)1).ToList() , 7 , ProjectID, ConvertedStartYear , ConvertedEndYear); //Quality


                return Json(new { DiciplineTypes, DiciplineValues }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
            }

            else 
            {
                return Json(new { DiciplineTypes, DiciplineValues }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
            }
        }

        [DontWrapResult]
        public async Task<ActionResult> AverageToGetAFC(int? ProjectID, string Year = "")
        {
            DateTime? ConvertedStartYear = null;
            DateTime? ConvertedEndYear = null;

            TimeSpan CivilResult = TimeSpan.FromSeconds(0);
            TimeSpan ElectricalResult = TimeSpan.FromSeconds(0);
            TimeSpan InstrumentResult = TimeSpan.FromSeconds(0);
            TimeSpan MechanicalResult = TimeSpan.FromSeconds(0);
            TimeSpan PipingResult = TimeSpan.FromSeconds(0);
            TimeSpan ProcessResult = TimeSpan.FromSeconds(0);
            TimeSpan QualityControlResult = TimeSpan.FromSeconds(0);

            string[] DiciplineTypes = { "Civil", "Electrical", "Instrument", "Mechanical", "Piping", "Process", "Quality" };

            if (Year != null)
            {
                var StartDate = Year + "/01/01";
                var EndDate = Year + "/12/29";

                ConvertedStartYear = PersianDateTime.Parse(StartDate);
                ConvertedEndYear = PersianDateTime.Parse(EndDate);
            }

            var DocumentsFilter = await _projectsTasksAppService.AverageToGetAFC(ProjectID, ConvertedStartYear, ConvertedEndYear);

            var DocumentsResult = DocumentsFilter.Select(a => a.Id);

            foreach(var Item in DocumentsResult)
            {
                var RevisionsFilter = await _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(Item);

                var TaskID = RevisionsFilter.Select(a => a.TaskID).FirstOrDefault();
                var FirstTRNo = RevisionsFilter.Select(a => a.TransmitalDate).FirstOrDefault();
                var LastTRNo = RevisionsFilter.Select(a => a.TransmitalDate).LastOrDefault();

                var Document = await _projectsTasksAppService.GetSpecificProjectsTask(TaskID);

                if (Document.Select(a => a.Dicipline).FirstOrDefault() == (ProjectsTasksDiciplineTypes)1) { TimeSpan Span = (TimeSpan)(LastTRNo - FirstTRNo); CivilResult = new TimeSpan(CivilResult.Ticks + Span.Ticks); };
                if (Document.Select(a => a.Dicipline).FirstOrDefault() == (ProjectsTasksDiciplineTypes)2) { TimeSpan Span = (TimeSpan)(LastTRNo - FirstTRNo); ElectricalResult = new TimeSpan(ElectricalResult.Ticks + Span.Ticks); };
                if (Document.Select(a => a.Dicipline).FirstOrDefault() == (ProjectsTasksDiciplineTypes)3) { TimeSpan Span = (TimeSpan)(LastTRNo - FirstTRNo); InstrumentResult = new TimeSpan(InstrumentResult.Ticks + Span.Ticks); };
                if (Document.Select(a => a.Dicipline).FirstOrDefault() == (ProjectsTasksDiciplineTypes)4) { TimeSpan Span = (TimeSpan)(LastTRNo - FirstTRNo); MechanicalResult = new TimeSpan(MechanicalResult.Ticks + Span.Ticks); };
                if (Document.Select(a => a.Dicipline).FirstOrDefault() == (ProjectsTasksDiciplineTypes)5) { TimeSpan Span = (TimeSpan)(LastTRNo - FirstTRNo); PipingResult = new TimeSpan(PipingResult.Ticks + Span.Ticks); };
                if (Document.Select(a => a.Dicipline).FirstOrDefault() == (ProjectsTasksDiciplineTypes)6) { TimeSpan Span = (TimeSpan)(LastTRNo - FirstTRNo); ProcessResult = new TimeSpan(ProcessResult.Ticks + Span.Ticks); };
                if (Document.Select(a => a.Dicipline).FirstOrDefault() == (ProjectsTasksDiciplineTypes)7) { TimeSpan Span = (TimeSpan)(LastTRNo - FirstTRNo); QualityControlResult = new TimeSpan(QualityControlResult.Ticks + Span.Ticks); };
            };

            //Get AFCRevisions Then Get Documents That Have AFCRevisions Then Specify Them By Dicipline Types

            var CountAFCCivil = DocumentsFilter.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)1).Count();
            var CountAFCElectrical = DocumentsFilter.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)2).Count();
            var CountAFCInstrument = DocumentsFilter.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)3).Count();
            var CountAFCMechanical = DocumentsFilter.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)4).Count();
            var CountAFCPiping = DocumentsFilter.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)5).Count();
            var CountAFCProcess = DocumentsFilter.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)6).Count();
            var CountAFCQualityControl = DocumentsFilter.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)7).Count();

            float[] DiciplineValues = { (float)CivilResult.Days / (float)CountAFCCivil, (float)ElectricalResult.Days / (float)CountAFCElectrical, (float)InstrumentResult.Days / (float)CountAFCInstrument, (float)MechanicalResult.Days / (float)CountAFCMechanical, (float)PipingResult.Days / (float)CountAFCPiping, (float)ProcessResult.Days / (float)CountAFCProcess , (float)QualityControlResult.Days / (float)CountAFCQualityControl };

            return Json(new { DiciplineTypes, DiciplineValues }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [DontWrapResult]
        public async Task<ActionResult> DiciplineProduceProcess(int? ProjectID = null, int? Dicipline = null, string StartYear = null , string StartMonth = null)
        {

            string StartDate = "";

            DateTime ConvertedStartDate = default(DateTime);

            List<string> ChartTitles = new List<string>();

            List<int> ChartValues = new List<int>();

            var CurrentDate = new PersianDateTime(DateTime.Now);



            ProjectsTasksViewModel.PersianDigitToEnglish PersianDigitToEnglish = delegate (string Parameter)
            {
                Dictionary<char, char> LettersDictionary = new Dictionary<char, char> { ['۰'] = '0', ['۱'] = '1', ['۲'] = '2', ['۳'] = '3', ['۴'] = '4', ['۵'] = '5', ['۶'] = '6', ['۷'] = '7', ['۸'] = '8', ['۹'] = '9', ['/'] = '/' };

                foreach (var item in Parameter)
                {
                    Parameter = Parameter.Replace(item, LettersDictionary[item]);
                }
                return Parameter.ToString();
            };

            ProjectsTasksViewModel.DiciplineProduceProcess DiciplineProduceProcessFunction = delegate (List<ProjectsTasksRevisions> FilterRevisionsList, int? ImportYear , int? ImportStartMonth , int? ImportEndMonth )
             {

                 for (int MonthCounter = (int)ImportStartMonth ; MonthCounter <= ImportEndMonth ; MonthCounter++)
                 {
                     var ForLoopStartDate = "";

                     var ForLoopEndDate = "";

                     if (MonthCounter == 12) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/29"; }
                     else if (MonthCounter == 11) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/30"; }
                     else if (MonthCounter == 10) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/30"; }
                     else if (MonthCounter == 9) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/30"; }
                     else if (MonthCounter == 8) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/30"; }
                     else if (MonthCounter == 7) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/30"; }
                     else if (MonthCounter == 6) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/31"; }
                     else if (MonthCounter == 5) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/31"; }
                     else if (MonthCounter == 4) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/31"; }
                     else if (MonthCounter == 3) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/31"; }
                     else if (MonthCounter == 2) { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/31"; }
                     else { ForLoopStartDate = ImportYear + "/" + MonthCounter + "/01"; ForLoopEndDate = ImportYear + "/" + MonthCounter + "/31"; }

                     var Result = FilterRevisionsList.Where(a => a.TransmitalDate >= PersianDateTime.Parse(ForLoopStartDate) && a.TransmitalDate <= PersianDateTime.Parse(ForLoopEndDate)).Count();

                     switch (MonthCounter)
                     {
                         case 1:
                             ChartTitles.Add("فروردین " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 2:
                             ChartTitles.Add("اردیبهشت " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 3:
                             ChartTitles.Add("خرداد " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 4:
                             ChartTitles.Add("تیر " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 5:
                             ChartTitles.Add("مرداد " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 6:
                             ChartTitles.Add("شهریور " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 7:
                             ChartTitles.Add("مهر " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 8:
                             ChartTitles.Add("آبان " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 9:
                             ChartTitles.Add("آذر " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 10:
                             ChartTitles.Add("دی " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 11:
                             ChartTitles.Add("بهمن " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         case 12:
                             ChartTitles.Add("اسفند " + ImportYear);
                             ChartValues.Add(Result);
                             break;
                         default:
                             break;
                     }
                 }
             };



            if (StartYear == null && StartMonth == null)
            {
                StartYear = PersianDigitToEnglish(new PersianDateTime(DateTime.Now.AddYears(-1)).ToString().Split("/")[0].ToString()); //It Returns One Year Before Now 

                StartMonth = "01";

                StartDate = StartYear + "/" + StartMonth + "/01";

                ConvertedStartDate = PersianDateTime.Parse(StartDate);
            }
            else if (StartYear != null && StartMonth != null)
            {
                if (StartMonth == "12" || StartMonth == "11" || StartMonth == "10") { StartDate = StartYear + "/" + StartMonth + "/01"; } else { StartDate = StartYear + "/0" + StartMonth + "/01"; };

                ConvertedStartDate = PersianDateTime.Parse(StartDate);
            }
            else if (StartYear != null && StartMonth == null)
            {
                StartDate = StartYear + "/01/01";

                ConvertedStartDate = PersianDateTime.Parse(StartDate);
            }



            var FilterDocuments = await _projectsTasksAppService.DiciplineProduceProcess(ProjectID, (ProjectsTasksDiciplineTypes?)Dicipline);

            var DocumentIDs = FilterDocuments.Select(a => a.Id).ToList();

            var FilterRevisions = await _projectsTasksRevisionsAppService.DiciplineProduceProcess(DocumentIDs, ConvertedStartDate);

            for (int YearCounter = Convert.ToInt32(new PersianDateTime(ConvertedStartDate).Year.ToString()) ; YearCounter <= Convert.ToInt32(new PersianDateTime(DateTime.Now).Year.ToString()) ; YearCounter++)
            {
                int ImportYear = 0;
                int ImportStartMonth = 0;
                int ImportEndMonth = 0;

                //If The Chosen Year And Year That We Are In Are The Same - OR - Checking The Last Year
                if (Convert.ToInt32(new PersianDateTime(ConvertedStartDate).Year.ToString()) == Convert.ToInt32(new PersianDateTime(DateTime.Now).Year.ToString()) || YearCounter == Convert.ToInt32(new PersianDateTime(DateTime.Now).Year.ToString()))
                {
                    ImportYear = YearCounter;
                    ImportStartMonth = 1;
                    ImportEndMonth = Convert.ToInt32(new PersianDateTime(DateTime.Now).Month.ToString());

                    DiciplineProduceProcessFunction(FilterRevisions, ImportYear, ImportStartMonth, ImportEndMonth);
                }
                //Checking The First Year 
                else if (YearCounter == Convert.ToInt32(new PersianDateTime(ConvertedStartDate).Year.ToString()))
                {
                    if (StartMonth == null)  { ImportYear = YearCounter; ImportStartMonth = 1; ImportEndMonth = 12; }
                    
                    else if (StartMonth.Length > 0) { ImportYear = YearCounter; ImportStartMonth = Convert.ToInt32(new PersianDateTime(ConvertedStartDate).Month.ToString()); ImportEndMonth = 12; }

                    DiciplineProduceProcessFunction(FilterRevisions, ImportYear , ImportStartMonth , ImportEndMonth);
                }
                //The Years Between First And Last Year
                else
                {
                    ImportYear = YearCounter;
                    ImportStartMonth = 1;
                    ImportEndMonth = 12;

                    DiciplineProduceProcessFunction(FilterRevisions, ImportYear, ImportStartMonth , ImportEndMonth);
                }
            }
            return Json(new { ChartTitles , ChartValues }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [DontWrapResult]
        public async Task<ActionResult> AverageProjectsDocuments(string StartDate = null, string EndDate = null) 
        {
            //var InputYear = PersianDateTime.Parse(Year);

            DateTime? ConvertedStartDate = default(DateTime?);
            DateTime? ConvertedEndDate = default(DateTime?);

            if (StartDate != null) { ConvertedStartDate = PersianDateTime.Parse(StartDate); }
            
            if (EndDate != null) { ConvertedEndDate = PersianDateTime.Parse(EndDate); }


            var ProjectsNames = new List<string>();

            var ProjectsAverageDocuments = new List<string>();


            var ProjectsIDs = await _projectAppService.GetAllProjects();

            var AllRevisions = await _projectsTasksRevisionsAppService.AllRevisionsInPeriodOfTime(ConvertedStartDate , ConvertedEndDate);

            foreach (var Item in ProjectsIDs)
            {
                var ProjectsDocumentsResult = await _projectsTasksAppService.GetDocumentsByProjectID(Item.Id);

                var ProjectDocumentsIDs = ProjectsDocumentsResult.Select(a => a.Id).ToList();


                var ProjectDocumentsRevisions = await _projectsTasksRevisionsAppService.AverageProjectsDocumentsProducedRevisions(ProjectDocumentsIDs, ConvertedStartDate, ConvertedEndDate);

                if (ProjectDocumentsRevisions.Count() > 0)
                {
                    ProjectsNames.Add(Item.Title);

                    float Calculation = (float)ProjectDocumentsRevisions.Count() / (float)AllRevisions.Count() * 100;

                    ProjectsAverageDocuments.Add( Calculation.ToString("0.00") );
                }
            }

         return Json(new { ProjectsNames , ProjectsAverageDocuments }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [DontWrapResult]
        public async Task<ActionResult> DocumentsStatusTable()
        {
            //var InputYear = PersianDateTime.Parse(Year);

            var ProjectsNames = new List<string>();

            var ProjectsAverageDocuments = new List<float>();


            var ProjectsIDs = await _projectAppService.GetAllProjects();

            var AllDocuments = await _projectsTasksAppService.GetAllProjectsTasks();

            foreach (var Item in ProjectsIDs)
            {
                var Info = await _projectsTasksAppService.GetDocumentsByProjectID(Item.Id);

                if (Info.Count() > 0)
                {
                    ProjectsNames.Add(Item.Title);

                    ProjectsAverageDocuments.Add(Info.Count() * 100 / AllDocuments.Count());
                }
            }

            return Json(new { ProjectsNames, ProjectsAverageDocuments }, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [DontWrapResult]
        public async Task<ActionResult> DocumentsStatus(int? ProjectID)
        {

            DashboardDocsStatusViewModel.ActionTypeCounter ActionCounter = delegate (List<ProjectsTasks> Input)
            {
                var List = new List<DashboardDocsStatusViewModel>();

                List.Add(new DashboardDocsStatusViewModel
                {
                    IssuedByMSV = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)5).Count(),

                    CommentedByFSTCO = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)3).Count(),

                    CommentedByIDOM = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)4).Count(),

                    ApprovedByFSTCO = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)1).Count(),

                    ApprovedByIDOM = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)2).Count(),

                    AppWithNotesByFSTCO = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)16).Count(),

                    AppWithNotesByIDOM = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)17).Count(),

                    NotIssued = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)6).Count(),

                    Delete = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)7).Count(),

                    Total = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)1 || a.LastAction == (ProjectsTasksActionTypes)2 || a.LastAction == (ProjectsTasksActionTypes)3 || a.LastAction == (ProjectsTasksActionTypes)4 || a.LastAction == (ProjectsTasksActionTypes)5 || a.LastAction == (ProjectsTasksActionTypes)6 || a.LastAction == (ProjectsTasksActionTypes)7 || a.LastAction == (ProjectsTasksActionTypes)16 || a.LastAction == (ProjectsTasksActionTypes)17).Count(),

                    TotalIssued = Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)1 || a.LastAction == (ProjectsTasksActionTypes)2 || a.LastAction == (ProjectsTasksActionTypes)3 || a.LastAction == (ProjectsTasksActionTypes)4 || a.LastAction == (ProjectsTasksActionTypes)5 || a.LastAction == (ProjectsTasksActionTypes)6 || a.LastAction == (ProjectsTasksActionTypes)7 || a.LastAction == (ProjectsTasksActionTypes)16 || a.LastAction == (ProjectsTasksActionTypes)17).Count() - Input.Where(a => a.LastAction == (ProjectsTasksActionTypes)6 || a.LastAction == (ProjectsTasksActionTypes)7).Count()
                });

                return List;
            };

            DashboardDocsStatusViewModel.DocumentTypeCounter DocumentTypeCounter = delegate (string Title , List<ProjectsTasks> Input)
            {
                var List = new List<DashboardDocsStatusViewModel>();


                var BasicDesign = ActionCounter(Input.Where(a => a.DocumentType == (ProjectsTasksDocumentTypes)11).ToList());

                List.Add(new DashboardDocsStatusViewModel { DiciplineType = Title , DocumentType = "Basic Design", IssuedByMSV = BasicDesign.Select(a => a.IssuedByMSV).FirstOrDefault() , CommentedByFSTCO = BasicDesign.Select(a => a.CommentedByFSTCO).FirstOrDefault(), CommentedByIDOM = BasicDesign.Select(a => a.CommentedByIDOM).FirstOrDefault(), ApprovedByFSTCO = BasicDesign.Select(a => a.ApprovedByFSTCO).FirstOrDefault(), ApprovedByIDOM = BasicDesign.Select(a => a.ApprovedByIDOM).FirstOrDefault(), AppWithNotesByFSTCO = BasicDesign.Select(a => a.AppWithNotesByFSTCO).FirstOrDefault(), AppWithNotesByIDOM = BasicDesign.Select(a => a.AppWithNotesByIDOM).FirstOrDefault(), NotIssued = BasicDesign.Select(a => a.NotIssued).FirstOrDefault(), Delete = BasicDesign.Select(a => a.Delete).FirstOrDefault() , Total = BasicDesign.Select(a => a.Total).FirstOrDefault() , TotalIssued = BasicDesign.Select(a => a.TotalIssued).FirstOrDefault() });

                var DetailDesign = ActionCounter(Input.Where(a => a.DocumentType == (ProjectsTasksDocumentTypes)12).ToList());

                List.Add(new DashboardDocsStatusViewModel { DiciplineType = Title, DocumentType = "Detail Design", IssuedByMSV = DetailDesign.Select(a => a.IssuedByMSV).FirstOrDefault(), CommentedByFSTCO = DetailDesign.Select(a => a.CommentedByFSTCO).FirstOrDefault(), CommentedByIDOM = DetailDesign.Select(a => a.CommentedByIDOM).FirstOrDefault(), ApprovedByFSTCO = DetailDesign.Select(a => a.ApprovedByFSTCO).FirstOrDefault(), ApprovedByIDOM = DetailDesign.Select(a => a.ApprovedByIDOM).FirstOrDefault(), AppWithNotesByFSTCO = DetailDesign.Select(a => a.AppWithNotesByFSTCO).FirstOrDefault(), AppWithNotesByIDOM = DetailDesign.Select(a => a.AppWithNotesByIDOM).FirstOrDefault() , NotIssued = DetailDesign.Select(a => a.NotIssued).FirstOrDefault(), Delete = DetailDesign.Select(a => a.Delete).FirstOrDefault() , Total = DetailDesign.Select(a => a.Total).FirstOrDefault() , TotalIssued = DetailDesign.Select(a => a.TotalIssued).FirstOrDefault() });

                var Procurement = ActionCounter(Input.Where(a => a.DocumentType == (ProjectsTasksDocumentTypes)13).ToList());

                List.Add(new DashboardDocsStatusViewModel { DiciplineType = Title, DocumentType = "Procurement Engineering", IssuedByMSV = Procurement.Select(a => a.IssuedByMSV).FirstOrDefault(), CommentedByFSTCO = Procurement.Select(a => a.CommentedByFSTCO).FirstOrDefault(), CommentedByIDOM = Procurement.Select(a => a.CommentedByIDOM).FirstOrDefault(), ApprovedByFSTCO = Procurement.Select(a => a.ApprovedByFSTCO).FirstOrDefault(), ApprovedByIDOM = Procurement.Select(a => a.ApprovedByIDOM).FirstOrDefault(), AppWithNotesByFSTCO = Procurement.Select(a => a.AppWithNotesByFSTCO).FirstOrDefault(), AppWithNotesByIDOM = Procurement.Select(a => a.AppWithNotesByIDOM).FirstOrDefault() , NotIssued = Procurement.Select(a => a.NotIssued).FirstOrDefault(), Delete = Procurement.Select(a => a.Delete).FirstOrDefault() , Total = Procurement.Select(a => a.Total).FirstOrDefault() , TotalIssued = Procurement.Select(a => a.TotalIssued).FirstOrDefault() });

                return List;
            };

            List<ProjectsTasks> AllDocuments = new List<ProjectsTasks>();

            if (!ProjectID.HasValue)
            {
                AllDocuments = await _projectsTasksAppService.GetAllProjectsTasks();
            }
            else
            {
                AllDocuments = await _projectsTasksAppService.GetDocumentsByProjectID((int)ProjectID);
            }

            var Process = AllDocuments.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)6).ToList();

            var Mechanical = AllDocuments.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)4).ToList();

            var CivilAndArchitecture = AllDocuments.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)1).ToList();

            var Electrical = AllDocuments.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)2).ToList();

            var InstrumentAndControl = AllDocuments.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)3).ToList();

            var Piping = AllDocuments.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)5).ToList();

            var QualityControl = AllDocuments.Where(a => a.Dicipline == (ProjectsTasksDiciplineTypes)7).ToList();

            var AllDicipline = AllDocuments.ToList();

            //All Dicipline


            List<DashboardDocsStatusViewModel>[] Result = { 
                DocumentTypeCounter("Process" , Process) ,
                DocumentTypeCounter("Mechanical" , Mechanical),
                DocumentTypeCounter("Civil & Arichitecture" , CivilAndArchitecture),
                DocumentTypeCounter("Electrical" , Electrical),
                DocumentTypeCounter("Instrument & Control" , InstrumentAndControl),
                DocumentTypeCounter("Piping" , Piping),
                DocumentTypeCounter("Quality Control" , QualityControl),
                DocumentTypeCounter("All Dicipline" , AllDicipline)
            };

            return Json(Result);
        }

        [DontWrapResult]
        public async Task<ActionResult> TotalDocumentsStatus(int? ProjectID)
        {
            List<ProjectsTasks> AllDocuments = new List<ProjectsTasks>();

            if (!ProjectID.HasValue)
            {
                AllDocuments = await _projectsTasksAppService.GetAllProjectsTasks();
            }
            else
            {
                AllDocuments = await _projectsTasksAppService.GetDocumentsByProjectID((int)ProjectID);
            }

            int[] Result =
            {
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)5).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)3).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)4).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)1).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)2).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)6).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)7).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)16).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)17).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)1 || a.LastAction == (ProjectsTasksActionTypes)2 || a.LastAction == (ProjectsTasksActionTypes)3 || a.LastAction == (ProjectsTasksActionTypes)4 || a.LastAction == (ProjectsTasksActionTypes)5 || a.LastAction == (ProjectsTasksActionTypes)6 || a.LastAction == (ProjectsTasksActionTypes)7 || a.LastAction == (ProjectsTasksActionTypes)16 || a.LastAction == (ProjectsTasksActionTypes)17).Count(),
                AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)1 || a.LastAction == (ProjectsTasksActionTypes)2 || a.LastAction == (ProjectsTasksActionTypes)3 || a.LastAction == (ProjectsTasksActionTypes)4 || a.LastAction == (ProjectsTasksActionTypes)5 || a.LastAction == (ProjectsTasksActionTypes)6 || a.LastAction == (ProjectsTasksActionTypes)7 || a.LastAction == (ProjectsTasksActionTypes)16 || a.LastAction == (ProjectsTasksActionTypes)17).Count() - AllDocuments.Where(a => a.LastAction == (ProjectsTasksActionTypes)6 || a.LastAction == (ProjectsTasksActionTypes)7).Count()
            };

            return Json(Result);
        }
    }
}
