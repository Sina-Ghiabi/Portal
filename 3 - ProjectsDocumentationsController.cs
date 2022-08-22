using Abp.Application.Services.Dto;
using Abp.AspNetCore.Mvc.Authorization;
using Abp.Web.Models;
using Dapna.MSVPortal.Controllers;
using Dapna.MSVPortal.Projects;
using Dapna.MSVPortal.ProjectsDocumentations;
using Dapna.MSVPortal.ProjectsDocumentations.Dto;
using Dapna.MSVPortal.Web.ViewModels;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using MD.PersianDateTime;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using Dapna.MSVPortal.Enums;
using Dapna.MSVPortal.Projects.Dto;
using System.Data.SqlClient;
using System.Globalization;
using Dapna.MSVPortal.Web.Views.Shared.Components.ProjectsTasksFilterColumns;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelDataReader;
using System.Data;
using Grpc.Core;
using Abp.Runtime.Validation;
using System.Threading;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Dapna.MSVPortal.Web.Controllers
{
    [AbpMvcAuthorize]

    public class ProjectsDocumentationsController : MSVPortalControllerBase
    {
        private readonly IHostingEnvironment _hostingEnvironment;
        private readonly IChosenColumnsAppService _chosenColumnsAppService;
        private readonly IProjectAppService _projectAppService;
        private readonly IProjectsTasksAppService _projectsTasksAppService;
        private readonly IProjectsTasksRevisionsAppService _projectsTasksRevisionsAppService;
        private readonly IProject_User_MappingAppService _project_User_MappingAppService;
        public ProjectsDocumentationsController(IChosenColumnsAppService chosenColumnsAppService, IProjectAppService projectAppService, IProjectsTasksAppService projectsTasksAppService, IProjectsTasksRevisionsAppService projectsTasksRevisionsAppService, IProject_User_MappingAppService project_User_MappingAppService, IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
            _chosenColumnsAppService = chosenColumnsAppService;
            _projectAppService = projectAppService;
            _projectsTasksAppService = projectsTasksAppService;
            _projectsTasksRevisionsAppService = projectsTasksRevisionsAppService;
            _project_User_MappingAppService = project_User_MappingAppService;
        }

        //------------------------------------------ Project's Tasks Methods ----------------------------------------------


        public IActionResult ProjectsTasks()
        {
            //Users ChosenColumns For ProjectsTasksTable => ProjectsTasksTable's ID Equals To 6 => TableID = 6 ;
            var ChosenColumns = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == 6);

            var Model = new ProjectsTasksViewModel();

            //If Before Reloading The ProjectsTasks Revisions Page Was Open It Will Open That Revision Page On Reload 

            if (TempData["ProjectCode"] != null)
            {
                Model.ProjectID = _projectAppService.GetProjectIDByProjectCode((string)TempData["ProjectCode"]).Result.Select(a => a.Id).FirstOrDefault();
                Model.ProjectCode = (string)TempData["ProjectCode"];
                Model.TaskID = (int)TempData["TaskID"];
                Model.DocumentNumber = (string)TempData["DocumentNumber"];
                Model.DocumentTitle = (string)TempData["DocumentTitle"];
            }

            if (ChosenColumns != null)
            {
                ViewBag.UserID = ChosenColumns.UserID;
                ViewBag.TableID = ChosenColumns.TableID;
                ViewBag.Column1 = ChosenColumns.Column1;
                ViewBag.Column2 = ChosenColumns.Column2;
                ViewBag.Column3 = ChosenColumns.Column3;
                ViewBag.Column4 = ChosenColumns.Column4;
                ViewBag.Column5 = ChosenColumns.Column5;
                ViewBag.Column6 = ChosenColumns.Column6;
                ViewBag.Column7 = ChosenColumns.Column7;
                ViewBag.Column8 = ChosenColumns.Column8;
                ViewBag.Column9 = ChosenColumns.Column9;
                ViewBag.Column10 = ChosenColumns.Column10;
                ViewBag.Column11 = ChosenColumns.Column11;
                ViewBag.Column12 = ChosenColumns.Column12;
                ViewBag.Column13 = ChosenColumns.Column13;
                ViewBag.Column14 = ChosenColumns.Column14;
                ViewBag.Column15 = ChosenColumns.Column15;
                ViewBag.Column16 = ChosenColumns.Column16;
                ViewBag.Column17 = ChosenColumns.Column17;
                ViewBag.Column18 = ChosenColumns.Column18;
                ViewBag.Column19 = ChosenColumns.Column19;
                ViewBag.Column20 = ChosenColumns.Column20;
                ViewBag.Column21 = ChosenColumns.Column21;
                ViewBag.Column22 = ChosenColumns.Column22;
                ViewBag.Column23 = ChosenColumns.Column23;
                ViewBag.Column24 = ChosenColumns.Column24;
                ViewBag.Column25 = ChosenColumns.Column25;
                ViewBag.Column26 = ChosenColumns.Column26;
                ViewBag.Column27 = ChosenColumns.Column27;
            }

            var Projects = _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value);

            ViewBag.ProjectId = Projects.Result.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            return View(Model);
        }

        public IActionResult ProjectsTasksEditChosenColumns(int TableID)
        {
            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == TableID);

            if (Record != null)
            {
                ChosenColumnsViewModel ChosenColumnsViewModel = new ChosenColumnsViewModel
                {
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Record.Column1,
                    Column2 = Record.Column2,
                    Column3 = Record.Column3,
                    Column4 = Record.Column4,
                    Column5 = Record.Column5,
                    Column6 = Record.Column6,
                    Column7 = Record.Column7,
                    Column8 = Record.Column8,
                    Column9 = Record.Column9,
                    Column10 = Record.Column10,
                    Column11 = Record.Column11,
                    Column12 = Record.Column12,
                    Column13 = Record.Column13,
                    Column14 = Record.Column14,
                    Column15 = Record.Column15,
                    Column16 = Record.Column16,
                    Column17 = Record.Column17,
                    Column18 = Record.Column18,
                    Column19 = Record.Column19,
                    Column20 = Record.Column20,
                    Column21 = Record.Column21,
                    Column22 = Record.Column22,
                    Column23 = Record.Column23,
                    Column24 = Record.Column24,
                    Column25 = Record.Column25,
                    Column26 = Record.Column26,
                    Column27 = Record.Column27
                };
                return View(ChosenColumnsViewModel);
            }
            return View();
        }

        [DontWrapResult]
        public async Task<ActionResult> ProjectsTasksTableData([DataSourceRequest] DataSourceRequest request)
        {
            List<ProjectsTasksViewModel> List = new List<ProjectsTasksViewModel>();

            IDictionary<int, string> ProjectsNames = new Dictionary<int, string>();

            if (HttpContext.Session.GetInt32("Flag") == 1)
            {
                var Result = await ProjectsDocumentsFilter(2,
                    HttpContext.Session.GetInt32("SearchType") != null ? HttpContext.Session.GetInt32("SearchType") : null,
                    HttpContext.Session.GetInt32("ProjectID") != null ? HttpContext.Session.GetInt32("ProjectID") : null,
                    HttpContext.Session.GetString("ProjectCode") != null ? HttpContext.Session.GetString("ProjectCode") : null,
                    HttpContext.Session.GetString("CompanyName") != null ? HttpContext.Session.GetString("CompanyName") : null,
                    HttpContext.Session.GetString("DocumentTitle") != null ? HttpContext.Session.GetString("DocumentTitle") : null,
                    HttpContext.Session.GetString("DocumentNumber") != null ? HttpContext.Session.GetString("DocumentNumber") : null,
                    HttpContext.Session.GetInt32("Dicipline") != null ? HttpContext.Session.GetInt32("Dicipline") : null,
                    HttpContext.Session.GetString("ResponsiblePerson") != null ? HttpContext.Session.GetString("ResponsiblePerson") : null,
                    HttpContext.Session.GetInt32("DocumentType") != null ? HttpContext.Session.GetInt32("DocumentType") : null,
                    HttpContext.Session.GetString("Description") != null ? HttpContext.Session.GetString("Description") : null,
                    HttpContext.Session.GetInt32("Progress") != null ? HttpContext.Session.GetInt32("Progress") : null,
                    HttpContext.Session.GetString("BaseLineStart") != null ? HttpContext.Session.GetString("BaseLineStart") : null,
                    HttpContext.Session.GetString("BaseLineFinished") != null ? HttpContext.Session.GetString("BaseLineFinished") : null,
                    HttpContext.Session.GetInt32("OriginalDuration") != null ? HttpContext.Session.GetInt32("OriginalDuration") : null,
                    HttpContext.Session.GetInt32("SourceOfItem") != null ? HttpContext.Session.GetInt32("SourceOfItem") : null,
                    HttpContext.Session.GetInt32("ManPower") != null ? HttpContext.Session.GetInt32("ManPower") : null,
                    HttpContext.Session.GetString("Critical") == "True" ? true : (bool?)null,
                    HttpContext.Session.GetString("TransmitalNumber") != null ? HttpContext.Session.GetString("TransmitalNumber") : null,
                    HttpContext.Session.GetString("TransmitalDate") != null ? HttpContext.Session.GetString("TransmitalDate") : null,
                    HttpContext.Session.GetString("Start-TransmitalDate") != null ? HttpContext.Session.GetString("Start-TransmitalDate") : null,
                    HttpContext.Session.GetString("End-TransmitalDate") != null ? HttpContext.Session.GetString("End-TransmitalDate") : null,
                    HttpContext.Session.GetString("CommentSheetNumber") != null ? HttpContext.Session.GetString("CommentSheetNumber") : null,
                    HttpContext.Session.GetString("CommentSheetDate") != null ? HttpContext.Session.GetString("CommentSheetDate") : null,
                    HttpContext.Session.GetString("ReplySheetNumber") != null ? HttpContext.Session.GetString("ReplySheetNumber") : null,
                    HttpContext.Session.GetString("ReplySheetDate") != null ? HttpContext.Session.GetString("ReplySheetDate") : null,
                    HttpContext.Session.GetString("RevisionNumber") != null ? HttpContext.Session.GetString("RevisionNumber") : null,
                    HttpContext.Session.GetInt32("Status") != null ? HttpContext.Session.GetInt32("Status") : null,
                    HttpContext.Session.GetInt32("Action") != null ? HttpContext.Session.GetInt32("Action") : null
                    );

                List = Result.Value;

                return Json(List , new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
            }

            var Projects = await _projectsTasksAppService.GetAllProjectsTasks();

            foreach (var Item in Projects)
            {
                var Project = await _projectAppService.GetProjectByID((int)Item.ProjectID);

                var ProjectName = Project.Select(a => a.Title).ToList();

                if (!ProjectsNames.ContainsKey(Item.ProjectID))
                {
                    ProjectsNames.Add(Item.ProjectID, string.Join("", ProjectName));
                }
            }
            List.AddRange(Projects.Select(Record => new ProjectsTasksViewModel
            {
                //ProjectsTask Fields
                TaskID = Record.Id,
                ProjectID = Record.ProjectID,
                ProjectName = ProjectsNames[(int)Record.ProjectID],
                ProjectCode = Record.ProjectCode,
                CompanyName = Record.CompanyName,
                DocumentTitle = Record.DocumentTitle,
                DocumentNumber = Record.DocumentNumber,
                Dicipline = Record.Dicipline,
                ResponsiblePerson = Record.ResponsiblePerson,
                DocumentType = Record.DocumentType,
                Description = Record.Description,

                //ProjectsTaskSchedule Fields
                WeightFactor = Record.WeightFactor,
                Progress = Record.Progress,
                BaseLineStart = Record.BaseLineStart != null ? new PersianDateTime(Record.BaseLineStart).ToShortDateString() : "",
                BaseLineFinished = Record.BaseLineFinished != null ? new PersianDateTime(Record.BaseLineFinished).ToShortDateString() : "",
                PlanStart = Record.PlanStart != null ? new PersianDateTime(Record.PlanStart).ToShortDateString() : "",
                PlanFinished = Record.PlanFinished != null ? new PersianDateTime(Record.PlanFinished).ToShortDateString() : "",
                ActualStart = Record.ActualStart != null ? new PersianDateTime(Record.ActualStart).ToShortDateString() : "",
                ActualFinished = Record.ActualFinished != null ? new PersianDateTime(Record.ActualFinished).ToShortDateString() : "",
                OriginalDuration = Record.OriginalDuration,
                SourceOfItem = Record.SourceOfItem,
                ManPower = Record.ManPower,
                Critical = (bool)Record.Critical,

                //Fields Below Will Be Filled From Last Revision 
                LastTransmitalNumber = Record.LastTransmitalNumber,
                LastTransmitalDate = Record.LastTransmitalDate != null ? new PersianDateTime(Record.LastTransmitalDate).ToShortDateString() : "",
                LastRevisionNumber = Record.LastRevisionNumber,
                LastStatus = Record.LastStatus,
                LastAction = Record.LastAction,
            }));

            var dsResult = List.ToDataSourceResult(request);
            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        [HttpPost]
        [DontWrapResult]
        public async Task<ActionResult> CountTransmital()
        {
            var Result = await ProjectsDocumentsFilter(3,
                   HttpContext.Session.GetInt32("SearchType") != null ? HttpContext.Session.GetInt32("SearchType") : null,
                   HttpContext.Session.GetInt32("ProjectID") != null ? HttpContext.Session.GetInt32("ProjectID") : null,
                   HttpContext.Session.GetString("ProjectCode") != null ? HttpContext.Session.GetString("ProjectCode") : null,
                   HttpContext.Session.GetString("CompanyName") != null ? HttpContext.Session.GetString("CompanyName") : null,
                   HttpContext.Session.GetString("DocumentTitle") != null ? HttpContext.Session.GetString("DocumentTitle") : null,
                   HttpContext.Session.GetString("DocumentNumber") != null ? HttpContext.Session.GetString("DocumentNumber") : null,
                   HttpContext.Session.GetInt32("Dicipline") != null ? HttpContext.Session.GetInt32("Dicipline") : null,
                   HttpContext.Session.GetString("ResponsiblePerson") != null ? HttpContext.Session.GetString("ResponsiblePerson") : null,
                   HttpContext.Session.GetInt32("DocumentType") != null ? HttpContext.Session.GetInt32("DocumentType") : null,
                   HttpContext.Session.GetString("Description") != null ? HttpContext.Session.GetString("Description") : null,
                   HttpContext.Session.GetInt32("Progress") != null ? HttpContext.Session.GetInt32("Progress") : null,
                   HttpContext.Session.GetString("BaseLineStart") != null ? HttpContext.Session.GetString("BaseLineStart") : null,
                   HttpContext.Session.GetString("BaseLineFinished") != null ? HttpContext.Session.GetString("BaseLineFinished") : null,
                   HttpContext.Session.GetInt32("OriginalDuration") != null ? HttpContext.Session.GetInt32("OriginalDuration") : null,
                   HttpContext.Session.GetInt32("SourceOfItem") != null ? HttpContext.Session.GetInt32("SourceOfItem") : null,
                   HttpContext.Session.GetInt32("ManPower") != null ? HttpContext.Session.GetInt32("ManPower") : null,
                   HttpContext.Session.GetString("Critical") == "True" ? true : (bool?)null,
                   HttpContext.Session.GetString("TransmitalNumber") != null ? HttpContext.Session.GetString("TransmitalNumber") : null,
                   HttpContext.Session.GetString("TransmitalDate") != null ? HttpContext.Session.GetString("TransmitalDate") : null,
                   HttpContext.Session.GetString("Start-TransmitalDate") != null ? HttpContext.Session.GetString("Start-TransmitalDate") : null,
                   HttpContext.Session.GetString("End-TransmitalDate") != null ? HttpContext.Session.GetString("End-TransmitalDate") : null,
                   HttpContext.Session.GetString("CommentSheetNumber") != null ? HttpContext.Session.GetString("CommentSheetNumber") : null,
                   HttpContext.Session.GetString("CommentSheetDate") != null ? HttpContext.Session.GetString("CommentSheetDate") : null,
                   HttpContext.Session.GetString("ReplySheetNumber") != null ? HttpContext.Session.GetString("ReplySheetNumber") : null,
                   HttpContext.Session.GetString("ReplySheetDate") != null ? HttpContext.Session.GetString("ReplySheetDate") : null,
                   HttpContext.Session.GetString("RevisionNumber") != null ? HttpContext.Session.GetString("RevisionNumber") : null,
                   HttpContext.Session.GetInt32("Status") != null ? HttpContext.Session.GetInt32("Status") : null,
                   HttpContext.Session.GetInt32("Action") != null ? HttpContext.Session.GetInt32("Action") : null
                   );

            var List = Result.Value;

            return Json(List);
        }

        //AddProjectstTasks Page
        public async Task<ActionResult> AddProjectsTasks()
        {
            var items = await _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value);

            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();
            return View();
        }
        public async Task<IActionResult> AddProjectsTasksFunction(ProjectsTasksViewModel Model)
        {
            //Get Projects Code From Projects Table 
            var Project = await _projectAppService.GetProjectByID(Model.ProjectID);

            var Item = new ProjectsTasksDto()
            {
                CreatorUserId = AbpSession.UserId.Value,
                ProjectID = Model.ProjectID,
                ProjectCode = Project.Select(a => a.Code).FirstOrDefault(),
                CompanyName = Model.CompanyName,
                DocumentTitle = Model.DocumentTitle,
                DocumentNumber = Model.DocumentNumber,
                Dicipline = Model.Dicipline,
                ResponsiblePerson = Model.ResponsiblePerson,
                DocumentType = Model.DocumentType,
                Description = Model.Description,
                Critical = false //We Dont Fill Critical In This Place But If We Dont Give False Value Here , Critical Value Would Be Null Instead Of True Or False Then Reading Data From Database Would Make Problem
            };
            await _projectsTasksAppService.Create(Item);
            return RedirectToAction("ProjectsTasks");
        }

        public async Task<IActionResult> ProjectsTaskEdit(int TaskID)
        {
            var items = await _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value);

            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            var Record = _projectsTasksAppService.GetSpecificProjectsTask(TaskID).Result.Find(e => e.Id == TaskID);

            TempData["ProjectID"] = Record.ProjectID;
            TempData["ProjectCode"] = Record.ProjectCode;
            TempData["TaskID"] = Record.Id;
            TempData["DocumentNumber"] = Record.DocumentNumber;
            TempData["DocumentTitle"] = Record.DocumentTitle;

            if (Record != null)
            {
                ProjectsTasksViewModel ProjectsTaskViewModel = new ProjectsTasksViewModel
                {
                    //ProjectsTask Fields
                    CompanyName = Record.CompanyName,
                    Dicipline = Record.Dicipline,
                    ResponsiblePerson = Record.ResponsiblePerson,
                    DocumentType = Record.DocumentType,
                    Description = Record.Description
                };
                return View(ProjectsTaskViewModel);
            }
            return View();
        }

        public async Task<IActionResult> ProjectsTaskEditFunction(ProjectsTasksViewModel Model)
        {
            List<object> Data = new List<object>();

            var SpecificRevisions = await _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(Model.TaskID);

            if (SpecificRevisions != null)
            {
                Data.Add(SpecificRevisions.Select(a => a.TransmitalNumber).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.TransmitalDate).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.RevisionNumber).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.Status).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.Action).LastOrDefault());
            }

            var Project = await _projectAppService.GetProjectByID(Model.ProjectID);
            var Record = await _projectsTasksAppService.Get(new EntityDto<int>(Model.TaskID));

            //ProjectsTasks Fields
            Record.ProjectID = Model.ProjectID;
            Record.ProjectCode = Project.Select(a => a.Code).FirstOrDefault();
            Record.CompanyName = Model.CompanyName;
            Record.DocumentTitle = Model.DocumentTitle;
            Record.DocumentNumber = Model.DocumentNumber;
            Record.Dicipline = Model.Dicipline;
            Record.ResponsiblePerson = Model.ResponsiblePerson;
            Record.DocumentType = Model.DocumentType;
            Record.Description = Model.Description;

            //Fields Below Will Be Filled From Last Revision 
            Record.LastTransmitalNumber = (string)Data[0];
            Record.LastTransmitalDate = (DateTime?)Data[1];
            Record.LastRevisionNumber = (string)Data[2];
            Record.LastStatus = (ProjectsTasksStatusTypes?)Data[3];
            Record.LastAction = (ProjectsTasksActionTypes?)Data[4];

            await _projectsTasksAppService.Update(Record);

            return RedirectToAction("ProjectsTasks" , Model);
        }

        public async Task<IActionResult> ProjectsTaskSchedule(int TaskID)
        {

            var items = await _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value);

            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title
            }).ToList();

            var Record = _projectsTasksAppService.GetSpecificProjectsTask(TaskID).Result.Find(e => e.Id == TaskID);

            if (Record != null)
            {
                ProjectsTasksViewModel ProjectsTaskViewModel = new ProjectsTasksViewModel
                {
                    //ProjectsSchedule Fields
                    TaskID = Record.Id, //This Is TaskID
                    WeightFactor = Record.WeightFactor,
                    Progress = Record.Progress,
                    BaseLineStart = Record.BaseLineStart != null ? new PersianDateTime(Record.BaseLineStart).ToShortDateString() : "",
                    BaseLineFinished = Record.BaseLineFinished != null ? new PersianDateTime(Record.BaseLineFinished).ToShortDateString() : "",
                    PlanStart = Record.PlanStart != null ? new PersianDateTime(Record.PlanStart).ToShortDateString() : "",
                    PlanFinished = Record.PlanFinished != null ? new PersianDateTime(Record.PlanFinished).ToShortDateString() : "",
                    ActualStart = Record.ActualStart != null ? new PersianDateTime(Record.ActualStart).ToShortDateString() : "",
                    ActualFinished = Record.ActualFinished != null ? new PersianDateTime(Record.ActualFinished).ToShortDateString() : "",
                    OriginalDuration = Record.OriginalDuration,
                    SourceOfItem = Record.SourceOfItem,
                    ManPower = Record.ManPower,
                    Critical = (bool)Record.Critical
                };
                return View(ProjectsTaskViewModel);
            }
            return View();
        }

        public async Task<IActionResult> ProjectsTaskScheduleFunction(ProjectsTasksViewModel Model)
        {
            List<object> Data = new List<object>();

            var SpecificRevisions = await _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(Model.TaskID);

            if (SpecificRevisions != null)
            {
                Data.Add(SpecificRevisions.Select(a => a.TransmitalNumber).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.TransmitalDate).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.RevisionNumber).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.Status).LastOrDefault());
                Data.Add(SpecificRevisions.Select(a => a.Action).LastOrDefault());
            }

            var Record = await _projectsTasksAppService.Get(new EntityDto<int>(Model.TaskID));

            //ProjectsSchedule Fields
            Record.WeightFactor = Model.WeightFactor;
            Record.Progress = Model.Progress;
            Record.BaseLineStart = Model.BaseLineStart != null ? PersianDateTime.Parse(Model.BaseLineStart) : default(DateTime?);
            Record.BaseLineFinished = Model.BaseLineFinished != null ? PersianDateTime.Parse(Model.BaseLineFinished) : default(DateTime?);
            Record.PlanStart = Model.PlanStart != null ? PersianDateTime.Parse(Model.PlanStart) : default(DateTime?);
            Record.PlanFinished = Model.PlanFinished != null ? PersianDateTime.Parse(Model.PlanFinished) : default(DateTime?);
            Record.ActualStart = Model.ActualStart != null ? PersianDateTime.Parse(Model.ActualStart) : default(DateTime?);
            Record.ActualFinished = Model.ActualFinished != null ? PersianDateTime.Parse(Model.ActualFinished) : default(DateTime?);
            Record.OriginalDuration = Model.OriginalDuration;
            Record.SourceOfItem = Model.SourceOfItem;
            Record.ManPower = Model.ManPower;
            Record.Critical = Model.Critical;

            //Fields Below Will Be Filled From Last Revision 
            Record.LastTransmitalNumber = (string)Data[0];
            Record.LastTransmitalDate = (DateTime?)Data[1];
            Record.LastRevisionNumber = (string)Data[2];
            Record.LastStatus = (ProjectsTasksStatusTypes?)Data[3];
            Record.LastAction = (ProjectsTasksActionTypes?)Data[4];

            await _projectsTasksAppService.Update(Record);

            return RedirectToAction("ProjectsTasks");
        }

        [DontWrapResult]
        public IActionResult ProjectsTaskDelete(int TaskID, [DataSourceRequest] DataSourceRequest request)
        {
            var Result = "";

            var CountRevisions = _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(TaskID).Result.Count();

            if (CountRevisions == 0) { _projectsTasksAppService.Delete(new EntityDto(TaskID)); Result = "1"; } else { Result = "0"; }

            return Json(new[] { Result }.ToDataSourceResult(request, ModelState));
        }

        //Update The LastTransmitalNumber,LastTransmitalDate,LastRevisionNumber,LastStatus,LastAction OF ProjectsTask After Deleting The Revision
        public async Task<IActionResult> ProjectsTasksTableUpdate(int TaskID)
        {
            var Project = _projectsTasksAppService.GetSpecificProjectsTask(TaskID).Result;

            var SpecificRevisions = _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(TaskID).Result;

            var Record = await _projectsTasksAppService.Get(new EntityDto<int>(TaskID));

            //Fields Below Will Be Filled From Last Revision 
            Record.LastTransmitalNumber = SpecificRevisions.Select(a => a.TransmitalNumber).LastOrDefault();
            Record.LastTransmitalDate = SpecificRevisions.Select(a => a.TransmitalDate).LastOrDefault();
            Record.LastRevisionNumber = SpecificRevisions.Select(a => a.RevisionNumber).LastOrDefault();
            Record.LastStatus = SpecificRevisions.Select(a => a.Status).LastOrDefault();
            Record.LastAction = SpecificRevisions.Select(a => a.Action).LastOrDefault();

            await _projectsTasksAppService.Update(Record);

            return Json(true, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        public async Task<IActionResult> ProjectsTasksFilterColumnsViewComponent()
        {
            var Flag = HttpContext.Session.GetInt32("Flag");

            var items = await _project_User_MappingAppService.GetUserProjects(AbpSession.UserId.Value);

            ViewBag.ProjectId = items.Select(a => new ProjectViewModel()
            {
                Id = a.ProjectId,
                Title = a.Project.Title,
            }).ToList();

            ViewBag.SearchType = HttpContext.Session.GetInt32("SearchType");

            if (Flag == 1)
            {
                var Model = new ProjectsTasksFilterColumnsViewModel();

                if (HttpContext.Session.GetInt32("ProjectID") != null) { Model.ProjectID = HttpContext.Session.GetInt32("ProjectID"); }

                if (HttpContext.Session.GetString("ProjectCode") != null) { Model.ProjectCode = HttpContext.Session.GetString("ProjectCode"); }

                if (HttpContext.Session.GetString("CompanyName") != null) { Model.CompanyName = HttpContext.Session.GetString("CompanyName"); }

                if (HttpContext.Session.GetString("DocumentTitle") != null) { Model.DocumentTitle = HttpContext.Session.GetString("DocumentTitle"); }

                if (HttpContext.Session.GetString("DocumentNumber") != null) { Model.DocumentNumber = HttpContext.Session.GetString("DocumentNumber"); }

                if (HttpContext.Session.GetInt32("Dicipline") != null) { Model.Dicipline = (ProjectsTasksDiciplineTypes?)HttpContext.Session.GetInt32("Dicipline"); }

                if (HttpContext.Session.GetString("ResponsiblePerson") != null) { Model.ResponsiblePerson = HttpContext.Session.GetString("ResponsiblePerson"); }

                if (HttpContext.Session.GetInt32("DocumentType") != null) { Model.DocumentType = (ProjectsTasksDocumentTypes?)HttpContext.Session.GetInt32("DocumentType"); }

                if (HttpContext.Session.GetString("Description") != null) { Model.Description = HttpContext.Session.GetString("Description"); }

                if (HttpContext.Session.GetInt32("Progress") != null) { Model.Progress = HttpContext.Session.GetInt32("Progress"); }

                if (HttpContext.Session.GetString("BaseLineStart") != null) { Model.BaseLineStart = HttpContext.Session.GetString("BaseLineStart"); }

                if (HttpContext.Session.GetString("BaseLineFinished") != null) { Model.BaseLineFinished = HttpContext.Session.GetString("BaseLineFinished"); }

                if (HttpContext.Session.GetInt32("OriginalDuration") != null) { Model.OriginalDuration = HttpContext.Session.GetInt32("OriginalDuration"); }

                if (HttpContext.Session.GetInt32("SourceOfItem") != null) { Model.SourceOfItem = (ProjectsTasksSourceOfItemTypes?)HttpContext.Session.GetInt32("SourceOfItem"); }

                if (HttpContext.Session.GetInt32("ManPower") != null) { Model.ManPower = HttpContext.Session.GetInt32("ManPower"); }

                if (HttpContext.Session.GetString("Critical") != null) { Model.Critical = HttpContext.Session.GetString("Critical") == "True" ? true : false; }

                if (HttpContext.Session.GetString("TransmitalNumber") != null) { Model.TransmitalNumber = HttpContext.Session.GetString("TransmitalNumber"); }

                if (HttpContext.Session.GetString("TransmitalDate") != null) { Model.TransmitalDate = HttpContext.Session.GetString("TransmitalDate"); }

                if (HttpContext.Session.GetString("Start-TransmitalDate") != null) { Model.StartTransmitalDate = HttpContext.Session.GetString("Start-TransmitalDate"); }

                if (HttpContext.Session.GetString("End-TransmitalDate") != null) { Model.EndTransmitalDate = HttpContext.Session.GetString("End-TransmitalDate"); }

                if (HttpContext.Session.GetString("CommentSheetNumber") != null) { Model.CommentSheetNumber = HttpContext.Session.GetString("CommentSheetNumber"); }

                if (HttpContext.Session.GetString("CommentSheetDate") != null) { Model.CommentSheetDate = HttpContext.Session.GetString("CommentSheetDate"); }

                if (HttpContext.Session.GetString("ReplySheetNumber") != null) { Model.ReplySheetNumber = HttpContext.Session.GetString("ReplySheetNumber"); }

                if (HttpContext.Session.GetString("ReplySheetDate") != null) { Model.ReplySheetDate = HttpContext.Session.GetString("ReplySheetDate"); }

                if (HttpContext.Session.GetString("RevisionNumber") != null) { Model.RevisionNumber = HttpContext.Session.GetString("RevisionNumber"); }

                if (HttpContext.Session.GetInt32("Status") != null) { Model.Status = HttpContext.Session.GetInt32("Status"); }

                if (HttpContext.Session.GetInt32("Action") != null) { Model.Action = HttpContext.Session.GetInt32("Action"); }

                return ViewComponent("ProjectsTasksFilterColumns", Model);
            }

            return ViewComponent("ProjectsTasksFilterColumns");
        }

        public async Task<ActionResult<List<ProjectsTasksViewModel>>> ProjectsDocumentsFilter(int ReturnType, int? SearchType = null, int? ProjectID = null, string ProjectCode = null, string CompanyName = null, string DocumentTitle = null, string DocumentNumber = null, int? Dicipline = null, string ResponsiblePerson = null, int? DocumentType = null, string Description = null, int? Progress = null, string BaseLineStart = null, string BaseLineFinished = null, int? OriginalDuration = null, int? SourceOfItem = null, int? ManPower = null, bool? Critical = null, string TransmitalNumber = null, string TransmitalDate = null, string StartTransmitalDate = null, string EndTransmitalDate = null, string CommentSheetNumber = null, string CommentSheetDate = null, string ReplySheetNumber = null, string ReplySheetDate = null, string RevisionNumber = null, int? Status = null, int? Action = null)
        {
            SetFilterColumnsSessions(SearchType, ProjectID, ProjectCode, CompanyName, DocumentTitle, DocumentNumber, (ProjectsTasksDiciplineTypes?)Dicipline, ResponsiblePerson, (ProjectsTasksDocumentTypes?)DocumentType, Description, Progress, BaseLineStart, BaseLineFinished, OriginalDuration, (ProjectsTasksSourceOfItemTypes?)SourceOfItem, ManPower, Critical, TransmitalNumber, TransmitalDate, CommentSheetNumber, CommentSheetDate, ReplySheetNumber, ReplySheetDate, RevisionNumber, (ProjectsTasksStatusTypes?)Status, (ProjectsTasksActionTypes?)Action, StartTransmitalDate, EndTransmitalDate);

            var DocumentsIDs = new List<int>();

            var DocumentsList = new List<ProjectsTasksViewModel>();

            var NumberOfTransmitals = 0;

            //Search In All Documents & All Revisions 

            if (SearchType == 1)
            {
                var RevisionsFilter = await _projectsTasksRevisionsAppService.RevisionsFilterResult(TransmitalNumber, TransmitalDate != null ? PersianDateTime.Parse(TransmitalDate) : default(DateTime?), StartTransmitalDate != null ? PersianDateTime.Parse(StartTransmitalDate) : default(DateTime?), EndTransmitalDate != null ? PersianDateTime.Parse(EndTransmitalDate) : default(DateTime?), CommentSheetNumber, CommentSheetDate != null ? PersianDateTime.Parse(CommentSheetDate) : default(DateTime?), ReplySheetNumber, ReplySheetDate != null ? PersianDateTime.Parse(ReplySheetDate) : default(DateTime?), RevisionNumber, (ProjectsTasksStatusTypes?)Status, (ProjectsTasksActionTypes?)Action);

                var TaskIDs = RevisionsFilter.Select(a => a.TaskID).ToList();

                var DocumentsResult = await _projectsTasksAppService.ListedTasksFilterResult(TaskIDs, ProjectID, ProjectCode, CompanyName, DocumentTitle, DocumentNumber, (ProjectsTasksDiciplineTypes?)Dicipline, ResponsiblePerson, (ProjectsTasksDocumentTypes?)DocumentType, Description, Progress, BaseLineStart != null ? PersianDateTime.Parse(BaseLineStart) : default(DateTime?), BaseLineFinished != null ? PersianDateTime.Parse(BaseLineFinished) : default(DateTime?), OriginalDuration, (ProjectsTasksSourceOfItemTypes?)SourceOfItem, ManPower, Critical);

                DocumentsIDs = DocumentsResult.Select(a => a.Id).ToList();
            }

            //Search In All Documents And Last Revisions

            else if (SearchType == 2)
            {
                var AllRevisions = await _projectsTasksRevisionsAppService.GetAllProjectsTaskRevisions();

                var LastRevisions = AllRevisions.GroupBy(a => a.TaskID, (Key, b) => b.OrderByDescending(a => a.TransmitalDate).First());

                var LastRevisionsIDs = LastRevisions.Select(a => a.Id).ToList();

                var LastRevisionsFilter = await _projectsTasksRevisionsAppService.LastRevisionsFilterResult(LastRevisionsIDs, TransmitalNumber, TransmitalDate != null ? PersianDateTime.Parse(TransmitalDate) : default(DateTime?), StartTransmitalDate != null ? PersianDateTime.Parse(StartTransmitalDate) : default(DateTime?), EndTransmitalDate != null ? PersianDateTime.Parse(EndTransmitalDate) : default(DateTime?), CommentSheetNumber, CommentSheetDate != null ? PersianDateTime.Parse(CommentSheetDate) : default(DateTime?), ReplySheetNumber, ReplySheetDate != null ? PersianDateTime.Parse(ReplySheetDate) : default(DateTime?), RevisionNumber, (ProjectsTasksStatusTypes?)Status, (ProjectsTasksActionTypes?)Action);

                var TaskIDs = LastRevisionsFilter.Select(a => a.TaskID).ToList();

                var DocumentsResult = await _projectsTasksAppService.ListedTasksFilterResult(TaskIDs, ProjectID, ProjectCode, CompanyName, DocumentTitle, DocumentNumber, (ProjectsTasksDiciplineTypes?)Dicipline, ResponsiblePerson, (ProjectsTasksDocumentTypes?)DocumentType, Description, Progress, BaseLineStart != null ? PersianDateTime.Parse(BaseLineStart) : default(DateTime?), BaseLineFinished != null ? PersianDateTime.Parse(BaseLineFinished) : default(DateTime?), OriginalDuration, (ProjectsTasksSourceOfItemTypes?)SourceOfItem, ManPower, Critical);

                DocumentsIDs = DocumentsResult.Select(a => a.Id).ToList();
            }

            //Search In All Documents

            else if (SearchType == 3)
            {
                var DocumentsResult = await _projectsTasksAppService.TasksFilterResult(ProjectID, ProjectCode, CompanyName, DocumentTitle, DocumentNumber, (ProjectsTasksDiciplineTypes?)Dicipline, ResponsiblePerson, (ProjectsTasksDocumentTypes?)DocumentType, Description, Progress, BaseLineStart != null ? PersianDateTime.Parse(BaseLineStart) : default(DateTime?), BaseLineFinished != null ? PersianDateTime.Parse(BaseLineFinished) : default(DateTime?), OriginalDuration, (ProjectsTasksSourceOfItemTypes?)SourceOfItem, ManPower, Critical);

                DocumentsIDs = DocumentsResult.Select(a => a.Id).ToList();
            }


            //Search In All Revisions

            else if (SearchType == 4)
            {
                var RevisionsFilter = await _projectsTasksRevisionsAppService.RevisionsFilterResult(TransmitalNumber, TransmitalDate != null ? PersianDateTime.Parse(TransmitalDate) : default(DateTime?), StartTransmitalDate != null ? PersianDateTime.Parse(StartTransmitalDate) : default(DateTime?), EndTransmitalDate != null ? PersianDateTime.Parse(EndTransmitalDate) : default(DateTime?), CommentSheetNumber, CommentSheetDate != null ? PersianDateTime.Parse(CommentSheetDate) : default(DateTime?), ReplySheetNumber, ReplySheetDate != null ? PersianDateTime.Parse(ReplySheetDate) : default(DateTime?), RevisionNumber, (ProjectsTasksStatusTypes?)Status, (ProjectsTasksActionTypes?)Action);

                DocumentsIDs = RevisionsFilter.Select(a => a.TaskID).ToList();
            }

            //Search In Last Revisions

            else if (SearchType == 5)
            {
                var AllRevisions = await _projectsTasksRevisionsAppService.GetAllProjectsTaskRevisions();

                var LastRevisions = AllRevisions.GroupBy(a => a.TaskID, (Key, b) => b.OrderByDescending(a => a.TransmitalDate).First());

                var LastRevisionsIDs = LastRevisions.Select(a => a.Id).ToList();

                var LastRevisionsFilter = await _projectsTasksRevisionsAppService.LastRevisionsFilterResult(LastRevisionsIDs, TransmitalNumber, TransmitalDate != null ? PersianDateTime.Parse(TransmitalDate) : default(DateTime?), StartTransmitalDate != null ? PersianDateTime.Parse(StartTransmitalDate) : default(DateTime?), EndTransmitalDate != null ? PersianDateTime.Parse(EndTransmitalDate) : default(DateTime?), CommentSheetNumber, CommentSheetDate != null ? PersianDateTime.Parse(CommentSheetDate) : default(DateTime?), ReplySheetNumber, ReplySheetDate != null ? PersianDateTime.Parse(ReplySheetDate) : default(DateTime?), RevisionNumber, (ProjectsTasksStatusTypes?)Status, (ProjectsTasksActionTypes?)Action);

                DocumentsIDs = LastRevisionsFilter.Select(a => a.TaskID).ToList();
            }

            //Search Document That Doesn't Have Revisions

            else if (SearchType == 6)
            {
                var DocumentsResult = await _projectsTasksAppService.TasksFilterResult(ProjectID, ProjectCode, CompanyName, DocumentTitle, DocumentNumber, (ProjectsTasksDiciplineTypes?)Dicipline, ResponsiblePerson, (ProjectsTasksDocumentTypes?)DocumentType, Description, Progress, BaseLineStart != null ? PersianDateTime.Parse(BaseLineStart) : default(DateTime?), BaseLineFinished != null ? PersianDateTime.Parse(BaseLineFinished) : default(DateTime?), OriginalDuration, (ProjectsTasksSourceOfItemTypes?)SourceOfItem, ManPower, Critical);

                var FilteredDocumentIDs = DocumentsResult.Select(a => a.Id).ToList();

                foreach (var ID in FilteredDocumentIDs)
                {
                    var CheckExistanse = await _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(ID);

                    var CountDocumentRevisions = CheckExistanse.Count();

                    if (CountDocumentRevisions == 0)
                    {
                        DocumentsIDs.Add(ID);
                    }
                }
            }

            else
            {
                var Result = await _projectsTasksAppService.GetAllProjectsTasks();

                DocumentsIDs = Result.Select(a => a.Id).ToList();
            }

            foreach (var ID in DocumentsIDs)
            {
                var Result = _projectsTasksAppService.GetSpecificProjectsTask(ID).Result.FirstOrDefault();

                DocumentsList.Add(new ProjectsTasksViewModel
                {
                    TaskID = Result.Id,
                    ProjectID = Result.ProjectID,
                    ProjectName = _projectAppService.GetProjectByID(Result.ProjectID).Result.Select(a => a.Title).FirstOrDefault(),
                    ProjectCode = Result.ProjectCode,
                    CompanyName = Result.CompanyName,
                    DocumentTitle = Result.DocumentTitle,
                    DocumentNumber = Result.DocumentNumber,
                    Dicipline = Result.Dicipline,
                    ResponsiblePerson = Result.ResponsiblePerson,
                    DocumentType = Result.DocumentType,
                    Description = Result.Description,
                    WeightFactor = Result.WeightFactor,
                    Progress = Result.Progress,
                    BaseLineStart = Result.BaseLineStart != null ? new PersianDateTime(Result.BaseLineStart).ToShortDateString() : "",
                    BaseLineFinished = Result.BaseLineFinished != null ? new PersianDateTime(Result.BaseLineFinished).ToShortDateString() : "",
                    PlanStart = Result.PlanStart != null ? new PersianDateTime(Result.PlanStart).ToShortDateString() : "",
                    PlanFinished = Result.PlanFinished != null ? new PersianDateTime(Result.PlanFinished).ToShortDateString() : "",
                    ActualStart = Result.ActualStart != null ? new PersianDateTime(Result.ActualStart).ToShortDateString() : "",
                    ActualFinished = Result.ActualFinished != null ? new PersianDateTime(Result.ActualFinished).ToShortDateString() : "",
                    OriginalDuration = Result.OriginalDuration,
                    SourceOfItem = Result.SourceOfItem,
                    ManPower = Result.ManPower,
                    Critical = (bool)Result.Critical,
                    LastTransmitalNumber = Result.LastTransmitalNumber,
                    LastTransmitalDate = Result.LastTransmitalDate != null ? new PersianDateTime(Result.LastTransmitalDate).ToShortDateString() : "",
                    LastRevisionNumber = Result.LastRevisionNumber,
                    LastStatus = Result.LastStatus,
                    LastAction = Result.LastAction
                });

                //Count Number Of Searched Transmitals

                var Transmitals = await _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(ID);

                NumberOfTransmitals = NumberOfTransmitals + Transmitals.Count();
            }

            //If ReturnType Equals 1 It Means Return Data As JSON For View 

            if (ReturnType == 1)
            {
                return Json(DocumentsList, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
            }

            //If ReturhType Equals 2 It Means Return Data As List For Backend

            else if (ReturnType == 2)
            {
                return DocumentsList;
            }

            //If ReturnType Equals 3 It Means Return Number Of Transmitals 

            else
            {
                var List = new List<ProjectsTasksViewModel>();

                List.Add(new ProjectsTasksViewModel { CountTransmitals = NumberOfTransmitals });

                return List;
            }

        }

        //---------------------------------- Project's Task's Revisions Methods ----------------------------------------------




        [DontWrapResult]
        public async Task<IActionResult> ProjectsTaskRevisionsTableData(int TaskID, [DataSourceRequest] DataSourceRequest request)
        {
            List<ProjectsTasksRevisionsViewModel> List = new List<ProjectsTasksRevisionsViewModel>();

            var Projects = await _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions(TaskID);

            List.AddRange(Projects.Select(Record => new ProjectsTasksRevisionsViewModel
            {
                //Identifiers
                RevisionID = Record.Id,
                TaskID = Record.TaskID,
                RevisionNumber = Record.RevisionNumber, //This Field Will Be Filled After TransmitalNumber Enters

                TransmitalNumber = Record.TransmitalNumber, //User Enters The Transmital Number Then The Fields Below Will Be Filled With Katibe's Database's Data
                TransmitalDate = Record.TransmitalDate != null ? new PersianDateTime(Record.TransmitalDate).ToShortDateString() : "",
                CommentSheetNumber = Record.CommentSheetNumber,
                CommentSheetDate = Record.CommentSheetDate != null ? new PersianDateTime(Record.CommentSheetDate).ToShortDateString() : "",
                ReplySheetNumber = Record.ReplySheetNumber,
                ReplySheetDate = Record.ReplySheetDate != null ? new PersianDateTime(Record.ReplySheetDate).ToShortDateString() : "",
                Status = Record.Status,
                Action = Record.Action
            }));

            var dsResult = List.ToDataSourceResult(request);
            return Json(dsResult, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        public IActionResult ProjectsTasksRevisionsEditChosenColumns(int TableID, string ProjectCode, int TaskID, string DocumentNumber, string DocumentTitle)
        {
            var Model = new ChosenColumnsViewModel()
            {
                TableID = TableID,
                ProjectCode = ProjectCode,
                TaskID = TaskID,
                DocumentNumber = DocumentNumber,
                DocumentTitle = DocumentTitle
            };

            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == TableID);

            if (Record != null)
            {
                ChosenColumnsViewModel ChosenColumnsViewModel = new ChosenColumnsViewModel
                {
                    ProjectCode = ProjectCode,
                    TaskID = TaskID,
                    DocumentNumber = DocumentNumber,
                    DocumentTitle = DocumentTitle,

                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Record.Column1,
                    Column2 = Record.Column2,
                    Column3 = Record.Column3,
                    Column4 = Record.Column4,
                    Column5 = Record.Column5,
                    Column6 = Record.Column6,
                    Column7 = Record.Column7,
                    Column8 = Record.Column8,
                    Column9 = Record.Column9,
                    Column10 = Record.Column10,
                    Column11 = Record.Column11
                };
                return View(ChosenColumnsViewModel);
            }
            return View(Model);
        }

        [DontWrapResult]
        public IActionResult ProjectsTaskRevisionsViewComponentTable(string ProjectCode, int TaskID, string DocumentNumber, string DocumentTitle)
        {
            //Users ChosenColumns For ProjectsTasksRevisionsTable => ProjectsTasksRevisionsTable's ID Equals To 7 => TableID = 7 ;
            var ChosenColumns = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == 7);

            if (ChosenColumns != null)
            {
                ViewBag.UserID = ChosenColumns.UserID;
                ViewBag.TableID = ChosenColumns.TableID;
                ViewBag.Column1 = ChosenColumns.Column1;
                ViewBag.Column2 = ChosenColumns.Column2;
                ViewBag.Column3 = ChosenColumns.Column3;
                ViewBag.Column4 = ChosenColumns.Column4;
                ViewBag.Column5 = ChosenColumns.Column5;
                ViewBag.Column6 = ChosenColumns.Column6;
                ViewBag.Column7 = ChosenColumns.Column7;
                ViewBag.Column8 = ChosenColumns.Column8;
                ViewBag.Column9 = ChosenColumns.Column9;
                ViewBag.Column10 = ChosenColumns.Column10;
                ViewBag.Column11 = ChosenColumns.Column11;

                return ViewComponent("ProjectsTasksRevisions", new { ProjectCode, TaskID, DocumentNumber, DocumentTitle });
            }
            return ViewComponent("ProjectsTasksRevisions", new { ProjectCode, TaskID, DocumentNumber, DocumentTitle });
        }

        public IActionResult AddProjectsTasksRevisions(string ProjectCode, int TaskID, string DocumentNumber, string DocumentTitle)
        {
            var Model = new ProjectsTasksRevisionsViewModel()
            {
                ProjectCode = ProjectCode,
                TaskID = TaskID,
                DocumentNumber = DocumentNumber,
                DocumentTitle = DocumentTitle
            };

            return View(Model);
        }

        public async Task<IActionResult> AddProjectsTasksRevisionsFunction(ProjectsTasksRevisionsViewModel Model)
        {
            //If User Submits The Form

            if (Model.FormType == "1")
            {
                var Item = new ProjectsTasksRevisionsDto()
                {
                    //Identifiers
                    TaskID = (int)Model.TaskID,
                    RevisionNumber = Model.RevisionNumber, //This Field Will Be Filled After TransmitalNumber Enters

                    TransmitalNumber = Model.TransmitalNumber, //User Enters The Transmital Number Then The Fields Below Will Be Filled With Katibe's Database's Data
                    TransmitalDate = Model.TransmitalDate != null ? PersianDateTime.Parse(Model.TransmitalDate) : default(DateTime?),
                    CommentSheetNumber = Model.CommentSheetNumber,
                    CommentSheetDate = Model.CommentSheetDate != null ? PersianDateTime.Parse(Model.CommentSheetDate) : default(DateTime?),
                    ReplySheetNumber = Model.ReplySheetNumber,
                    ReplySheetDate = Model.ReplySheetDate != null ? PersianDateTime.Parse(Model.ReplySheetDate) : default(DateTime?),
                    Status = Model.Status,
                    Action = Model.Action
                };
                await _projectsTasksRevisionsAppService.Create(Item);

                //Update LastTransmitalNumber,LastTransmitalDate,LastRevisionNumber,LastStatus,LastAction Field In ProjectsTasks Table From Last Revision Added For That Task
                var Record = await _projectsTasksAppService.Get(new EntityDto<int>((int)Model.TaskID));

                //Fields Below Will Be Filled From Last Revision
                Record.LastTransmitalNumber = Model.TransmitalNumber;
                Record.LastTransmitalDate = Model.TransmitalDate != null ? PersianDateTime.Parse(Model.TransmitalDate) : default(DateTime?);
                Record.LastRevisionNumber = Model.RevisionNumber;
                Record.LastStatus = Model.Status;
                Record.LastAction = Model.Action;

                await _projectsTasksAppService.Update(Record);

                TempData["ProjectCode"] = Model.ProjectCode;
                TempData["TaskID"] = Model.TaskID;
                TempData["DocumentNumber"] = Model.DocumentNumber;
                TempData["DocumentTitle"] = Model.DocumentTitle;

                return RedirectToAction("ProjectsTasks");
            }

            //If User Canceled 
            else
            {
                TempData["ProjectCode"] = Model.ProjectCode;
                TempData["TaskID"] = Model.TaskID;
                TempData["DocumentNumber"] = Model.DocumentNumber;
                TempData["DocumentTitle"] = Model.DocumentTitle;

                return RedirectToAction("ProjectsTasks");
            }
        }


        public IActionResult AddProjectsTasksRevisionsFromKatibeh(string ProjectCode, int TaskID, string DocumentNumber, string DocumentTitle)
        {
            var Model = new ProjectsTasksRevisionsViewModel()
            {
                ProjectCode = ProjectCode,
                TaskID = TaskID,
                DocumentNumber = DocumentNumber,
                DocumentTitle = DocumentTitle
            };

            return View(Model);
        }

        [DontWrapResult]
        public IActionResult ReadFromKatibeh(string ProjectCode, string DocumentNumber, string TransmitalNumber)
        {
            List<ProjectsTasksRevisionsViewModel> List = new List<ProjectsTasksRevisionsViewModel>();

            ProjectsTasksViewModel.ConvertAction ConvertAction = delegate (string Parameter)
            {
                int Value = 0;

                Parameter.Replace(" ", "");

                switch (Parameter)
                {
                    case "15599":
                        Value = 1;
                        break;
                    case "15624":
                        Value = 2;
                        break;
                    case "15674":
                        Value = 3;
                        break;
                    case "15699":
                        Value = 4;
                        break;
                    case "15724":
                        Value = 4;
                        break;
                    case "15774":
                        Value = 5;
                        break;
                    case "15799":
                        Value = 6;
                        break;
                    case "15824":
                        Value = 7;
                        break;
                    case "13974":
                        Value = 8;
                        break;
                    case "14024":
                        Value = 9;
                        break;
                    case "11874":
                        Value = 10;
                        break;
                    case "13949":
                        Value = 11;
                        break;
                    case "11899":
                        Value = 12;
                        break;
                    case "13924":
                        Value = 13;
                        break;
                    case "13999":
                        Value = 14;
                        break;
                    case "14549":
                        Value = 15;
                        break;
                    case "15874":
                        Value = 16;
                        break;
                    case "15849":
                        Value = 17;
                        break;
                }
                return Value;
            };

            ProjectsTasksViewModel.EnglishDigitToPersian EnglishDigitToPersian = delegate (string Parameter)
            {
                Dictionary<char, char> LettersDictionary = new Dictionary<char, char>
                {
                    ['0'] = '۰',
                    ['1'] = '۱',
                    ['2'] = '۲',
                    ['3'] = '۳',
                    ['4'] = '۴',
                    ['5'] = '۵',
                    ['6'] = '۶',
                    ['7'] = '۷',
                    ['8'] = '۸',
                    ['9'] = '۹',
                    ['/'] = '/'
                };
                foreach (var item in Parameter)
                {
                    Parameter = Parameter.Replace(item, LettersDictionary[item]);
                }
                return Parameter.ToString();
            };

            ProjectsTasksViewModel.DateConverter DateConverter = delegate (string Parameter)
            {
                var Result = "";
                if (Parameter == null) { Result = ""; }
                else { Result = new PersianDateTime(DateTime.Parse(Parameter)).ToShortDateString(); }
                return Result;
            };

            SqlConnection SQLConnection = new SqlConnection("Server=epms.msv.net\\SQLSERVER2014; Database=FTV_EO; user id=varjavand; password=123456789; Trusted_Connection=false;");
            var Query = "SELECT mdr_SysCode AS[System Code] , On_LettCode AS[Project Code] , mdr_Number AS[Document Number] , [Trans No] AS[Transmital Number] , [Trans Date] AS[Transmital Date], (SELECT TOP  1 TaNameh.NameCode AS [Comment Sheet Number] FROM DCC_TransDetail INNER JOIN DCC_TransLett ON DCC_TransLett.tl_TrSysCode = DCC_TransDetail.trd_Master LEFT JOIN DCC_Trans ON DCC_Trans.trn_SysCode = DCC_TransDetail.trd_Master INNER JOIN TaNameh ON (TaNameh.SOmran = DCC_TransLett.tl_LettSysCode AND TaNameh.NoeName = DCC_TransLett.tl_LettType) WHERE TaNameh.NoeName = 1 AND DCC_Trans.trn_SysCode IN (SELECT trn_SysCode FROM dbo.DCC_TransDetail JOIN dbo.DCC_Mdr ON mdr_SysCode=trd_MdrSysCode JOIN dbo.DCC_Trans ON trd_Master=trn_SysCode WHERE mdr_PrjSysCode = (SELECT On_SysCode FROM TaOnvanTree WHERE On_LettCode = '" + ProjectCode + "') AND mdr_Number = '" + DocumentNumber + "' AND trn_TrNumber LIKE '%" + TransmitalNumber + "%') AND trn_TrNumber NOT LIKE '%IN%' ORDER BY [Comment Sheet Number] DESC) AS [Comment Sheet Number], (SELECt TOP  1 dbo.InsertSlashTarikh(TaNameh.TarikhSabt) AS [Comment Sheet Date] FROM DCC_TransDetail INNER JOIN DCC_TransLett ON DCC_TransLett.tl_TrSysCode = DCC_TransDetail.trd_Master LEFT JOIN DCC_Trans ON DCC_Trans.trn_SysCode = DCC_TransDetail.trd_Master INNER JOIN TaNameh ON (TaNameh.SOmran = DCC_TransLett.tl_LettSysCode AND TaNameh.NoeName = DCC_TransLett.tl_LettType) WHERE trn_TrNumber Not like '%IN%' AND TaNameh.NoeName = 1 AND DCC_Trans.trn_SysCode IN (SELECT trn_SysCode FROM dbo.DCC_TransDetail JOIN dbo.DCC_Mdr ON mdr_SysCode=trd_MdrSysCode JOIN dbo.DCC_Trans ON trd_Master=trn_SysCode WHERE mdr_PrjSysCode = (SELECT On_SysCode FROM TaOnvanTree WHERE On_LettCode = '" + ProjectCode + "') AND mdr_Number = '" + DocumentNumber + "' AND trn_TrNumber LIKE '%" + TransmitalNumber + "%') AND trn_TrNumber NOT LIKE '%IN%' ORDER BY [Comment Sheet Date] DESC) AS [Comment Sheet Date] , (SELECt TOP  1 TaNameh.NameCode AS [Reply Sheet Number] FROM DCC_TransDetail INNER JOIN DCC_TransLett ON DCC_TransLett.tl_TrSysCode = DCC_TransDetail.trd_Master LEFT JOIN DCC_Trans ON DCC_Trans.trn_SysCode = DCC_TransDetail.trd_Master INNER JOIN TaNameh ON (TaNameh.SOmran = DCC_TransLett.tl_LettSysCode AND TaNameh.NoeName = DCC_TransLett.tl_LettType) WHERE TaNameh.NoeName = 0 AND DCC_Trans.trn_SysCode IN (SELECT trn_SysCode FROM dbo.DCC_TransDetail JOIN dbo.DCC_Mdr ON mdr_SysCode=trd_MdrSysCode JOIN dbo.DCC_Trans ON trd_Master=trn_SysCode WHERE mdr_PrjSysCode = (SELECT On_SysCode FROM TaOnvanTree WHERE On_LettCode = '" + ProjectCode + "') AND mdr_Number = '" + DocumentNumber + "' AND trn_TrNumber LIKE '%" + TransmitalNumber + "%') AND trn_TrNumber NOT LIKE '%IN%' ORDER BY [Reply Sheet Number] DESC) AS[Reply Sheet Number], (SELECt TOP  1 dbo.InsertSlashTarikh(TaNameh.TarikhSabt) AS [Reply Sheet Date] FROM DCC_TransDetail INNER JOIN DCC_TransLett ON DCC_TransLett.tl_TrSysCode = DCC_TransDetail.trd_Master LEFT JOIN DCC_Trans ON DCC_Trans.trn_SysCode = DCC_TransDetail.trd_Master INNER JOIN TaNameh ON (TaNameh.SOmran = DCC_TransLett.tl_LettSysCode AND TaNameh.NoeName = DCC_TransLett.tl_LettType) WHERE trn_TrNumber Not like '%IN%' AND TaNameh.NoeName = 0 AND DCC_Trans.trn_SysCode IN (SELECT trn_SysCode FROM dbo.DCC_TransDetail JOIN dbo.DCC_Mdr ON mdr_SysCode=trd_MdrSysCode JOIN dbo.DCC_Trans ON trd_Master=trn_SysCode WHERE mdr_PrjSysCode = (SELECT On_SysCode FROM TaOnvanTree WHERE On_LettCode = '" + ProjectCode + "') AND mdr_Number = '" + DocumentNumber + "' AND trn_TrNumber LIKE '%" + TransmitalNumber + "%') AND trn_TrNumber NOT LIKE '%IN%' ORDER BY [Reply Sheet Date] DESC) AS[Reply Sheet Date] , Revision AS [Revision Number] , StatusCaption AS [Status] , (SELECT  BID_SysCode FROM dbo.DCC_TransDetail JOIN dbo.DCC_Mdr ON mdr_SysCode=trd_MdrSysCode JOIN dbo.DCC_Trans ON trd_Master=trn_SysCode JOIN TaBaseInfoDetail ON trn_Status = TaBaseInfoDetail.BID_SysCode WHERE mdr_SysCode IN (SELECT mdr_SysCode from DCC_Mdr WHERE mdr_PrjSysCode IN (SELECT On_SysCode FROM TaOnvanTree WHERE On_LettCode = '" + ProjectCode + "') AND mdr_Number LIKE '%" + DocumentNumber + "%') AND trn_TrNumber LIKE '%" + TransmitalNumber + "%' AND trn_TrNumber Not Like '%IN%') AS [Action] FROM EO_VMdrRevLst JOIN TaOnvanTree on EO_VMdrRevLst.mdr_PrjSysCode = TaOnvanTree.On_SysCode WHERE mdr_PrjSysCode IN (SELECT On_SysCode FROM TaOnvanTree WHERE On_LettCode = '" + ProjectCode + "') and mdr_Number = '" + DocumentNumber + "' and [Trans No] like '%" + TransmitalNumber + "%' and [Trans No] NOT LIKE '%IN%'";
            SqlCommand Command = new SqlCommand(Query, SQLConnection);
            SQLConnection.Open();
            var Reader = Command.ExecuteReader();

            while (Reader.Read())
            {
                List.Add(new ProjectsTasksRevisionsViewModel { TransmitalNumber = Reader[3].ToString(), TransmitalDate = DateConverter(Reader[4].ToString()), CommentSheetNumber = Reader[5].ToString(), CommentSheetDate = EnglishDigitToPersian(Reader[6].ToString()), ReplySheetNumber = Reader[7].ToString(), ReplySheetDate = EnglishDigitToPersian(Reader[8].ToString()), RevisionNumber = Reader[9].ToString(), Status = !Reader.IsDBNull(10) ? (ProjectsTasksStatusTypes?)Enum.Parse(typeof(ProjectsTasksStatusTypes), (string)Reader[10]) : null, Action = !Reader.IsDBNull(11) ? (ProjectsTasksActionTypes?)Enum.Parse(typeof(ProjectsTasksActionTypes), ConvertAction(Reader[11].ToString()).ToString()) : null });
            }

            SQLConnection.Close();

            return Json(List, new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
        }

        public async Task<IActionResult> ChosenRevisionsFromKatibeh(string FormType, string ProjectCode, int TaskID, string DocumentNumber, string DocumentTitle, string TransmitalNumber, string TransmitalDate, string CommentSheetNumber, string CommentSheetDate, string ReplySheetNumber, string ReplySheetDate, string RevisionNumber, int? Status, int? Action)
        {
            if (FormType == "1")
            {
                var Result = await _projectsTasksRevisionsAppService.CheckRevisionsExistance(TaskID, TransmitalNumber);

                if (Result.Select(a => a.TransmitalNumber).FirstOrDefault() == null)
                {
                    var Item = new ProjectsTasksRevisionsDto()
                    {
                        //Identifiers
                        TaskID = TaskID,
                        RevisionNumber = RevisionNumber, //This Field Will Be Filled After TransmitalNumber Enters

                        TransmitalNumber = TransmitalNumber, //User Enters The Transmital Number Then The Fields Below Will Be Filled With Katibe's Database's Data
                        TransmitalDate = TransmitalDate != null ? PersianDateTime.Parse(TransmitalDate) : default(DateTime?),
                        CommentSheetNumber = CommentSheetNumber,
                        CommentSheetDate = CommentSheetDate != null ? PersianDateTime.Parse(CommentSheetDate) : default(DateTime?),
                        ReplySheetNumber = ReplySheetNumber,
                        ReplySheetDate = ReplySheetDate != null ? PersianDateTime.Parse(ReplySheetDate) : default(DateTime?),
                        Status = Status != null ? (ProjectsTasksStatusTypes?)Status : null,
                        Action = Action != null ? (ProjectsTasksActionTypes?)Action : null
                    };
                    await _projectsTasksRevisionsAppService.Create(Item);

                    //Update LastTransmitalNumber,LastTransmitalDate,LastRevisionNumber,LastStatus,LastAction Field In ProjectsTasks Table From Last Revision Added For That Task
                    var Record = await _projectsTasksAppService.Get(new EntityDto<int>(TaskID));

                    //Fields Below Will Be Filled From Last Revision
                    Record.LastTransmitalNumber = TransmitalNumber;
                    Record.LastTransmitalDate = TransmitalDate != null ? PersianDateTime.Parse(TransmitalDate) : default(DateTime?); ;
                    Record.LastRevisionNumber = RevisionNumber;
                    Record.LastStatus = Status != null ? (ProjectsTasksStatusTypes?)Status : null;
                    Record.LastAction = Action != null ? (ProjectsTasksActionTypes?)Action : null;

                    await _projectsTasksAppService.Update(Record);

                    TempData["ProjectCode"] = ProjectCode;
                    TempData["TaskID"] = TaskID;
                    TempData["DocumentNumber"] = DocumentNumber;
                    TempData["DocumentTitle"] = DocumentTitle;


                    //Return 1 Means , That Transmital Number Added And It Should Return To ProjectsTasks

                    return Json("1", new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
                }

                //Return 0 Means , That Transmital Number Existed In Database And It Is Not Allowed To Insert It 

                return Json("0", new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
            }
            else
            {
                TempData["ProjectCode"] = ProjectCode;
                TempData["TaskID"] = TaskID;
                TempData["DocumentNumber"] = DocumentNumber;
                TempData["DocumentTitle"] = DocumentTitle;

                //Return 1 Means That It Should Return To ProjectsTasks

                return Json("1", new JsonSerializerSettings() { ContractResolver = new DefaultContractResolver() });
            }
        }

        public IActionResult ProjectsTasksRevisionEdit(int RevisionID, int TaskID, string ProjectCode, string DocumentNumber, string DocumentTitle)
        {
            var Record = _projectsTasksRevisionsAppService.GetSpecificProjectsTasksRevision(RevisionID).Result.Find(e => e.Id == RevisionID);

            if (Record != null)
            {
                ProjectsTasksRevisionsViewModel ProjectsTaskRevisionsViewModel = new ProjectsTasksRevisionsViewModel
                {
                    //Identifiers
                    RevisionID = Record.Id,
                    TaskID = TaskID,
                    ProjectCode = ProjectCode,
                    DocumentNumber = DocumentNumber,
                    DocumentTitle = DocumentTitle,
                    RevisionNumber = Record.RevisionNumber, //This Field Will Be Filled After TransmitalNumber Enters

                    TransmitalNumber = Record.TransmitalNumber, //User Enters The Transmital Number Then The Fields Below Will Be Filled With Katibe's Database's Data
                    TransmitalDate = Record.TransmitalDate != null ? new PersianDateTime(Record.TransmitalDate).ToShortDateString() : "",
                    CommentSheetNumber = Record.CommentSheetNumber,
                    CommentSheetDate = Record.CommentSheetDate != null ? new PersianDateTime(Record.CommentSheetDate).ToShortDateString() : "",
                    ReplySheetNumber = Record.ReplySheetNumber,
                    ReplySheetDate = Record.ReplySheetDate != null ? new PersianDateTime(Record.ReplySheetDate).ToShortDateString() : "",
                    Status = Record.Status,
                    Action = Record.Action
                };
                return View(ProjectsTaskRevisionsViewModel);
            }
            return View();
        }

        public async Task<IActionResult> ProjectsTasksRevisionEditFunction(ProjectsTasksRevisionsViewModel Model)
        {
            if (Model.FormType == "1")
            {
                var RevisionRecord = await _projectsTasksRevisionsAppService.Get(new EntityDto<int>((int)Model.RevisionID));

                //Identifiers
                RevisionRecord.RevisionNumber = Model.RevisionNumber; //This Field Will Be Filled After TransmitalNumber Enters

                RevisionRecord.TransmitalNumber = Model.TransmitalNumber; //User Enters The Transmital Number Then The Fields Below Will Be Filled With Katibe's Database's Data
                RevisionRecord.TransmitalDate = Model.TransmitalDate != null ? PersianDateTime.Parse(Model.TransmitalDate) : default(DateTime?);
                RevisionRecord.CommentSheetNumber = Model.CommentSheetNumber;
                RevisionRecord.CommentSheetDate = Model.CommentSheetDate != null ? PersianDateTime.Parse(Model.CommentSheetDate) : default(DateTime?);
                RevisionRecord.ReplySheetNumber = Model.ReplySheetNumber;
                RevisionRecord.ReplySheetDate = Model.ReplySheetDate != null ? PersianDateTime.Parse(Model.ReplySheetDate) : default(DateTime?);
                RevisionRecord.Status = Model.Status;
                RevisionRecord.Action = Model.Action;

                await _projectsTasksRevisionsAppService.Update(RevisionRecord);

                //After Edition Update The ProjectsTask LastAction Field With Last Revisions LastAction 
                var Record = await _projectsTasksAppService.Get(new EntityDto<int>((int)Model.TaskID));
                var SpecificRevisions = _projectsTasksRevisionsAppService.GetSpecificProjectsTaskRevisions((int)Model.TaskID).Result;

                //Fields Below Will Be Filled From Last Revision
                Record.LastTransmitalNumber = SpecificRevisions.Select(a => a.TransmitalNumber).LastOrDefault();
                Record.LastTransmitalDate = SpecificRevisions.Select(a => a.TransmitalDate).LastOrDefault();
                Record.LastRevisionNumber = SpecificRevisions.Select(a => a.RevisionNumber).LastOrDefault();
                Record.LastStatus = SpecificRevisions.Select(a => a.Status).LastOrDefault();
                Record.LastAction = SpecificRevisions.Select(a => a.Action).LastOrDefault();

                await _projectsTasksAppService.Update(Record);

                TempData["ProjectCode"] = Model.ProjectCode;
                TempData["TaskID"] = Model.TaskID;
                TempData["DocumentNumber"] = Model.DocumentNumber;
                TempData["DocumentTitle"] = Model.DocumentTitle;

                return RedirectToAction("ProjectsTasks");
            }

            else
            {
                TempData["ProjectCode"] = Model.ProjectCode;
                TempData["TaskID"] = Model.TaskID;
                TempData["DocumentNumber"] = Model.DocumentNumber;
                TempData["DocumentTitle"] = Model.DocumentTitle;

                return RedirectToAction("ProjectsTasks");
            }

        }

        [AcceptVerbs("Post")]
        [DontWrapResult]
        [DisableValidation]
        public async Task<IActionResult> ProjectsTaskRevisionsDestroy([DataSourceRequest] DataSourceRequest request, ProjectsTasksRevisionsViewModel Item)
        {
            if (Item != null)
            {
                await _projectsTasksRevisionsAppService.Delete(new EntityDto((int)Item.RevisionID));
            }
            return Json(new[] { true }.ToDataSourceResult(request, ModelState));
        }


        //------------------------------------------------------ Other Methods ----------------------------------------------------------------
        public void SetFilterColumnsSessions(int? SearchType, int? ProjectID, string ProjectCode, string CompanyName, string DocumentTitle, string DocumentNumber, ProjectsTasksDiciplineTypes? Dicipline, string ResponsiblePerson, ProjectsTasksDocumentTypes? DocumentType, string Description, int? Progress, string BaseLineStart, string BaseLineFinished, int? OriginalDuration, ProjectsTasksSourceOfItemTypes? SourceOfItem, int? ManPower, bool? Critical, string TransmitalNumber, string TransmitalDate, string CommentSheetNumber, string CommentSheetDate, string ReplySheetNumber, string ReplySheetDate, string RevisionNumber, ProjectsTasksStatusTypes? Status, ProjectsTasksActionTypes? Action, string StartTransmitalDate, string EndTransmitalDate)
        {
            HttpContext.Session.SetInt32("Flag", 1);

            if (SearchType.HasValue) { HttpContext.Session.SetInt32("SearchType", (int)SearchType); } else { HttpContext.Session.Remove("SearchType"); }

            if (ProjectID.HasValue) { HttpContext.Session.SetInt32("ProjectID", (int)ProjectID); } else { HttpContext.Session.Remove("ProjectID"); }

            if (ProjectCode != null) { HttpContext.Session.SetString("ProjectCode", ProjectCode); } else { HttpContext.Session.Remove("ProjectCode"); }

            if (CompanyName != null) { HttpContext.Session.SetString("CompanyName", CompanyName); } else { HttpContext.Session.Remove("CompanyName"); }

            if (DocumentTitle != null) { HttpContext.Session.SetString("DocumentTitle", DocumentTitle); } else { HttpContext.Session.Remove("DocumentTitle"); }

            if (DocumentNumber != null) { HttpContext.Session.SetString("DocumentNumber", DocumentNumber); } else { HttpContext.Session.Remove("DocumentNumber"); }

            if (Dicipline.HasValue) { HttpContext.Session.SetInt32("Dicipline", (int)Dicipline); } else { HttpContext.Session.Remove("Dicipline"); }

            if (ResponsiblePerson != null) { HttpContext.Session.SetString("ResponsiblePerson", ResponsiblePerson); } else { HttpContext.Session.Remove("ResponsiblePerson"); }

            if (DocumentType.HasValue) { HttpContext.Session.SetInt32("DocumentType", (int)DocumentType); } else { HttpContext.Session.Remove("DocumentType"); }

            if (Description != null) { HttpContext.Session.SetString("Description", Description); } else { HttpContext.Session.Remove("Description"); }

            if (Progress.HasValue) { HttpContext.Session.SetInt32("Progress", (int)Progress); } else { HttpContext.Session.Remove("Progress"); }

            if (BaseLineStart != null) { HttpContext.Session.SetString("BaseLineStart", BaseLineStart); } else { HttpContext.Session.Remove("BaseLineStart"); }

            if (BaseLineFinished != null) { HttpContext.Session.SetString("BaseLineFinished", BaseLineFinished); } else { HttpContext.Session.Remove("BaseLineFinished"); }

            if (OriginalDuration.HasValue) { HttpContext.Session.SetInt32("OriginalDuration", (int)OriginalDuration); } else { HttpContext.Session.Remove("OriginalDuration"); }

            if (SourceOfItem.HasValue) { HttpContext.Session.SetInt32("SourceOfItem", (int)SourceOfItem); } else { HttpContext.Session.Remove("SourceOfItem"); }

            if (ManPower.HasValue) { HttpContext.Session.SetInt32("ManPower", (int)ManPower); } else { HttpContext.Session.Remove("ManPower"); }

            if (Critical != null) { HttpContext.Session.SetString("Critical", Critical == true ? "True" : null); } else { HttpContext.Session.Remove("Critical"); }

            if (TransmitalNumber != null) { HttpContext.Session.SetString("TransmitalNumber", TransmitalNumber); } else { HttpContext.Session.Remove("TransmitalNumber"); }

            if (TransmitalDate != null) { HttpContext.Session.SetString("TransmitalDate", TransmitalDate); } else { HttpContext.Session.Remove("TransmitalDate"); }

            if (StartTransmitalDate != null) { HttpContext.Session.SetString("Start-TransmitalDate", StartTransmitalDate); } else { HttpContext.Session.Remove("Start-TransmitalDate"); }

            if (EndTransmitalDate != null) { HttpContext.Session.SetString("End-TransmitalDate", EndTransmitalDate); } else { HttpContext.Session.Remove("End-TransmitalDate"); }

            if (CommentSheetNumber != null) { HttpContext.Session.SetString("CommentSheetNumber", CommentSheetNumber); } else { HttpContext.Session.Remove("CommentSheetNumber"); }

            if (CommentSheetDate != null) { HttpContext.Session.SetString("CommentSheetDate", CommentSheetDate); } else { HttpContext.Session.Remove("CommentSheetDate"); }

            if (ReplySheetNumber != null) { HttpContext.Session.SetString("ReplySheetNumber", ReplySheetNumber); } else { HttpContext.Session.Remove("ReplySheetNumber"); }

            if (ReplySheetDate != null) { HttpContext.Session.SetString("ReplySheetDate", ReplySheetDate); } else { HttpContext.Session.Remove("ReplySheetDate"); }

            if (RevisionNumber != null) { HttpContext.Session.SetString("RevisionNumber", RevisionNumber); } else { HttpContext.Session.Remove("RevisionNumber"); }

            if (Status.HasValue) { HttpContext.Session.SetInt32("Status", (int)Status); } else { HttpContext.Session.Remove("Status"); }

            if (Action.HasValue) { HttpContext.Session.SetInt32("Action", (int)Action); } else { HttpContext.Session.Remove("Action"); }
        }

        public void DeleteFilterColumns()
        {
            HttpContext.Session.Remove("Flag");

            HttpContext.Session.Remove("SearchType");

            HttpContext.Session.Remove("ProjectID");

            HttpContext.Session.Remove("ProjectCode");

            HttpContext.Session.Remove("CompanyName");

            HttpContext.Session.Remove("DocumentTitle");

            HttpContext.Session.Remove("DocumentNumber");

            HttpContext.Session.Remove("Dicipline");

            HttpContext.Session.Remove("ResponsiblePerson");

            HttpContext.Session.Remove("DocumentType");

            HttpContext.Session.Remove("Description");

            HttpContext.Session.Remove("Progress");

            HttpContext.Session.Remove("BaseLineStart");

            HttpContext.Session.Remove("BaseLineFinished");

            HttpContext.Session.Remove("OriginalDuration");

            HttpContext.Session.Remove("SourceOfItem");

            HttpContext.Session.Remove("ManPower");

            HttpContext.Session.Remove("Critical");

            HttpContext.Session.Remove("TransmitalNumber");

            HttpContext.Session.Remove("TransmitalDate");

            HttpContext.Session.Remove("Start-TransmitalDate");

            HttpContext.Session.Remove("End-TransmitalDate");

            HttpContext.Session.Remove("CommentSheetNumber");

            HttpContext.Session.Remove("CommentSheetDate");

            HttpContext.Session.Remove("ReplySheetNumber");

            HttpContext.Session.Remove("ReplySheetDate");

            HttpContext.Session.Remove("RevisionNumber");

            HttpContext.Session.Remove("Status");

            HttpContext.Session.Remove("Action");
        }

        public async Task<IActionResult> AddOrUpdateChosenColumns(ChosenColumnsViewModel Model)
        {
            //If User Was In RevisionsChosenColumn Page And Canceled Changes This Function Must Be Triggred

            if (Model.FormType == "0")
            {
                TempData["ProjectCode"] = Model.ProjectCode;
                TempData["TaskID"] = Model.TaskID;
                TempData["DocumentNumber"] = Model.DocumentNumber;
                TempData["DocumentTitle"] = Model.DocumentTitle;

                return RedirectToAction("ProjectsTasks");
            }

            string TableSection = "";

            //GetChosenColumns Returns All The Users Tables , And Then We Sholud Check If The Tables ID Exists

            var Record = _chosenColumnsAppService.GetChosenColumns(AbpSession.UserId.Value).Result.Find(e => e.TableID == Model.TableID);

            if (Model.TableID == 6)
            {
                TableSection = "/Views/ProjectsDocumentations/ProjectsTasks";
            }
            else if (Model.TableID == 7)
            {
                TableSection = "/Views/Shared/Components/ProjectsTasksRevisions/Default.cshtml";
            }

            if (Record == null)
            {
                await _chosenColumnsAppService.Create(new ChosenColumnsDto()
                {
                    UserID = (int)AbpSession.UserId.Value,
                    Section = TableSection,
                    TableID = Model.TableID,
                    Column1 = Model.Column1,
                    Column2 = Model.Column2,
                    Column3 = Model.Column3,
                    Column4 = Model.Column4,
                    Column5 = Model.Column5,
                    Column6 = Model.Column6,
                    Column7 = Model.Column7,
                    Column8 = Model.Column8,
                    Column9 = Model.Column9,
                    Column10 = Model.Column10,
                    Column11 = Model.Column11,
                    Column12 = Model.Column12,
                    Column13 = Model.Column13,
                    Column14 = Model.Column14,
                    Column15 = Model.Column15,
                    Column16 = Model.Column16,
                    Column17 = Model.Column17,
                    Column18 = Model.Column18,
                    Column19 = Model.Column19,
                    Column20 = Model.Column20,
                    Column21 = Model.Column21,
                    Column22 = Model.Column22,
                    Column23 = Model.Column23,
                    Column24 = Model.Column24,
                    Column25 = Model.Column25,
                    Column26 = Model.Column26,
                    Column27 = Model.Column27
                });
            }
            else
            {
                await _chosenColumnsAppService.Update(new ChosenColumnsDto
                {
                    Id = Record.Id,
                    UserID = Record.UserID,
                    TableID = Record.TableID,
                    Column1 = Model.Column1,
                    Column2 = Model.Column2,
                    Column3 = Model.Column3,
                    Column4 = Model.Column4,
                    Column5 = Model.Column5,
                    Column6 = Model.Column6,
                    Column7 = Model.Column7,
                    Column8 = Model.Column8,
                    Column9 = Model.Column9,
                    Column10 = Model.Column10,
                    Column11 = Model.Column11,
                    Column12 = Model.Column12,
                    Column13 = Model.Column13,
                    Column14 = Model.Column14,
                    Column15 = Model.Column15,
                    Column16 = Model.Column16,
                    Column17 = Model.Column17,
                    Column18 = Model.Column18,
                    Column19 = Model.Column19,
                    Column20 = Model.Column20,
                    Column21 = Model.Column21,
                    Column22 = Model.Column22,
                    Column23 = Model.Column23,
                    Column24 = Model.Column24,
                    Column25 = Model.Column25,
                    Column26 = Model.Column26,
                    Column27 = Model.Column27,
                });
            }

            if (Model.TableID == 6)
            {
                return RedirectToAction("ProjectsTasks");
            }

            //If User Submited The Revisions ChosenColumn Form This Function Must Be Triggred
            if (Model.FormType == "1")
            {
                TempData["ProjectCode"] = Model.ProjectCode;
                TempData["TaskID"] = Model.TaskID;
                TempData["DocumentNumber"] = Model.DocumentNumber;
                TempData["DocumentTitle"] = Model.DocumentTitle;

                return RedirectToAction("ProjectsTasks");
            }
            else
            {
                return RedirectToAction("ProjectsTasks");
            }
        }

        public IActionResult ReadExcel()
        {
            return View();
        }

        public async Task<IActionResult> ImportExcel(IFormFile ExcelFile)
        {
            ProjectsTasksViewModel.ConvertToGregorian ConvertToGregorian = delegate (string Parameter)
            {
                string[] Date = Parameter.Split('/');
                PersianCalendar PersianCalendar = new PersianCalendar();
                string ConvertResult = new DateTime(int.Parse(Date[0]), int.Parse(Date[1]), int.Parse(Date[2]), PersianCalendar).ToString("yyyy-MM-dd HH:mm:ss.fffffff");
                return ConvertResult;
            };


            ProjectsTasksViewModel.ConvertAction ConvertAction = delegate (string Parameter)
            {
                int Value = 0;

                Parameter.Replace(" ", "");

                switch (Parameter)
                {
                    case "Approved By FSTCO":
                        Value = 1;
                        break;
                    case "Approved By IDOM":
                        Value = 2;
                        break;
                    case "Commented By FSTCO":
                        Value = 3;
                        break;
                    case "Commented By IDOM":
                        Value = 4;
                        break;
                    case "Comments By IDOM":
                        Value = 4;
                        break;
                    case "Issued By MSV":
                        Value = 5;
                        break;
                    case "Not Issue":
                        Value = 6;
                        break;
                    case "Delete":
                        Value = 7;
                        break;
                    case "تأييد شده":
                        Value = 8;
                        break;
                    case "تأييد شده مشروط":
                        Value = 9;
                        break;
                    case "در حال بررسي ماشين سازي ويژه":
                        Value = 10;
                        break;
                    case "در حال بررسي فكور صنعت":
                        Value = 11;
                        break;
                    case "در حال بررسي كارفرما":
                        Value = 12;
                        break;
                    case "در حال بررسي وندور":
                        Value = 13;
                        break;
                    case "حذف شده":
                        Value = 14;
                        break;
                    case "-----":
                        Value = 15;
                        break;
                    case "App With Notes By IDOM":
                        Value = 16;
                        break;
                    case "App With Notes By FST":
                        Value = 17;
                        break;
                    default:
                        break;
                }
                return Value;
            };


            IDictionary<int, string> ExistedTasks = new Dictionary<int, string>();
            IDictionary<int, string> ErroredTasks = new Dictionary<int, string>();

            string FilePath = "";

            //var ExistedTasks = new object();
            //var DataProblem = new object();

            SqlConnection SQLConnection = new SqlConnection("Server=192.168.2.47; Database=msv_portal; user id=msvportal_user; password=Msv8810;Trusted_Connection=false");

            SQLConnection.Open();

            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel\\ProjectsTasks-ImportExcel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel\\ProjectsTasks-ImportExcel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            if (ExcelFile != null)
            {
                //Save The ExcelFile In Excel Folder In WWWRoot Path

                string Excels = Path.Combine(_hostingEnvironment.WebRootPath, "Excel");
                string ImportExcels = Path.Combine(Excels, "ProjectsTasks-ImportExcel");
                if (ExcelFile.FileName.Contains(".xlsx")) { FilePath = Path.Combine(ImportExcels, ExcelFile.FileName.Replace(".xlsx", " - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToLongDateTimeInt() + ".xlsx")); } //Add DateTime To Imported Excel File
                else if (ExcelFile.FileName.Contains(".xls")) { FilePath = Path.Combine(ImportExcels, ExcelFile.FileName.Replace(".xls", " - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToLongDateTimeInt() + ".xls")); } //Add DateTime To Imported Excel File
                else { return Json("فایل مورد نظر از نوع اکسل نمی باشد"); } // If File Wasn't Excel Format , Returns Error Alert

                using (Stream FileStream = new FileStream(FilePath, FileMode.Create)) { await ExcelFile.CopyToAsync(FileStream); } //Copy Imported Excel File In Excel Folder In WWWRoot 
            }


            FileStream Stream = System.IO.File.Open(FilePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader ExcelReader;


            //1.Reading Excel file
            if (Path.GetExtension(FilePath).ToUpper() == ".XLS")
            {
                //1.1 Reading from a binary Excel file('97-2003 format; *.xls)
                ExcelReader = ExcelReaderFactory.CreateBinaryReader(Stream);
            }
            else
            {
                //1.2 Reading from a OpenXml Excel file(2007 format; *.xlsx)
                ExcelReader = ExcelReaderFactory.CreateOpenXmlReader(Stream);
            }

            //2.DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet Result = ExcelReader.AsDataSet();
            DataTable DataTable = Result.Tables[0];

            var CountRows = DataTable.Rows.Count; //Number Of All The Rows In Excel 

            for (int Index = 1; Index < CountRows; Index++)
            {
                var CheckExistence = _projectsTasksAppService.CheckTaskExistence(Convert.ToInt32(DataTable.Rows[Index][0]), Convert.ToString(DataTable.Rows[Index][3])).Result.Select(a => a.Id).ToList();

                //If We Didn't Have Same ProjectID & DocumentNumber Or DocumentNumber Was Null Insert The Row
                if (CheckExistence.Count() == 0)
                {
                    try
                    {
                        var ProjectID = Convert.ToInt32(DataTable.Rows[Index][0]);
                        var ProjectCode = Convert.ToString(DataTable.Rows[Index][1]);
                        var CompanyName = Convert.ToString(DataTable.Rows[Index][2]);
                        var DocumentTitle = Convert.ToString(DataTable.Rows[Index][3]);
                        var DocumentNumber = !DataTable.Rows[Index].IsNull(4) ? Convert.ToString(DataTable.Rows[Index][4]) : null;
                        var Dicipline = Convert.ToInt32((ProjectsTasksDiciplineTypes)Enum.Parse(typeof(ProjectsTasksDiciplineTypes), Convert.ToString(DataTable.Rows[Index][5]).Replace("&", "And").Replace(" ", "")));
                        var ResponsiblePerson = !DataTable.Rows[Index].IsNull(6) ? Convert.ToString(DataTable.Rows[Index][6]) : null;
                        var DocumentType = !DataTable.Rows[Index].IsNull(7) ? Convert.ToInt32((ProjectsTasksDocumentTypes?)Enum.Parse(typeof(ProjectsTasksDocumentTypes), Convert.ToString(DataTable.Rows[Index][7]).Replace("&", "And").Replace(" ", ""))) : (int?)null;
                        var Description = !DataTable.Rows[Index].IsNull(8) ? Convert.ToString(DataTable.Rows[Index][8]) : null;
                        var Critical = !DataTable.Rows[Index].IsNull(9) ? (bool?)DataTable.Rows[Index][9] : false;
                        var WeightFactor = !DataTable.Rows[Index].IsNull(10) ? Convert.ToInt32(DataTable.Rows[Index][10]) : (int?)null;
                        var Progress = !DataTable.Rows[Index].IsNull(11) ? Convert.ToInt32(DataTable.Rows[Index][11]) : (int?)null;
                        var BaseLineStart = !DataTable.Rows[Index].IsNull(12) ? ConvertToGregorian(Convert.ToString(DataTable.Rows[Index][12])) : null;
                        var BaseLineFinished = !DataTable.Rows[Index].IsNull(13) ? ConvertToGregorian(Convert.ToString(DataTable.Rows[Index][13])) : null;
                        var PlanStart = !DataTable.Rows[Index].IsNull(14) ? ConvertToGregorian(Convert.ToString(DataTable.Rows[Index][14])) : null;
                        var PlanFinished = !DataTable.Rows[Index].IsNull(15) ? ConvertToGregorian(Convert.ToString(DataTable.Rows[Index][15])) : null;
                        var ActualStart = !DataTable.Rows[Index].IsNull(16) ? ConvertToGregorian(Convert.ToString(DataTable.Rows[Index][16])) : null;
                        var ActualFinished = !DataTable.Rows[Index].IsNull(17) ? ConvertToGregorian(Convert.ToString(DataTable.Rows[Index][17])) : null;
                        var OriginalDuration = !DataTable.Rows[Index].IsNull(18) ? Convert.ToInt32(DataTable.Rows[Index][18].ToString().Replace("d", "")) : (int?)null;
                        var SourceOfItem = !DataTable.Rows[Index].IsNull(19) ? Convert.ToInt32((ProjectsTasksSourceOfItemTypes?)Enum.Parse(typeof(ProjectsTasksSourceOfItemTypes), Convert.ToString(DataTable.Rows[Index][19]).Replace(" ", ""))) : (int?)null;
                        var ManPower = !DataTable.Rows[Index].IsNull(20) ? Convert.ToInt32(DataTable.Rows[Index][20].ToString()) : (int?)null;
                        var LastTransmitalNumber = !DataTable.Rows[Index].IsNull(21) ? Convert.ToString(DataTable.Rows[Index][21]) : null;
                        var LastTransmitalDate = !DataTable.Rows[Index].IsNull(22) ? PersianDateTime.Parse(Convert.ToString(DataTable.Rows[Index][22])).ToString() : null;
                        var LastRevisionNumber = !DataTable.Rows[Index].IsNull(23) ? Convert.ToString(DataTable.Rows[Index][23]) : null;
                        var LastStatus = !DataTable.Rows[Index].IsNull(24) ? Convert.ToInt32((ProjectsTasksStatusTypes?)Enum.Parse(typeof(ProjectsTasksStatusTypes), Convert.ToString(DataTable.Rows[Index][24]).Replace(" ", ""))) : (int?)null;
                        var LastAction = !DataTable.Rows[Index].IsNull(25) ? ConvertAction(Convert.ToString(DataTable.Rows[Index][25])) : (int?)null;

                        var Query = "INSERT INTO ProjectsTasks(CreationTime, CreatorUserId , IsDeleted , ProjectID , ProjectCode , CompanyName , DocumentTitle , DocumentNumber , Dicipline , ResponsiblePerson , DocumentType , Description , Critical , WeightFactor , Progress , BaseLineStart , BaseLineFinished , PlanStart , PlanFinished , ActualStart , ActualFinished , OriginalDuration , SourceOfItem , ManPower , LastTransmitalNumber , LastTransmitalDate , LastRevisionNumber , LastStatus , LastAction) VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffff") + "', " + AbpSession.UserId.Value + " , 'FALSE' , " + ProjectID + " , '" + ProjectCode + "' , '" + CompanyName + "' , '" + DocumentTitle + "' , '" + DocumentNumber + "' , " + Dicipline + " , N'" + ResponsiblePerson + "' , " + DocumentType + " , N'" + Description + "' , '" + Critical + "' , " + WeightFactor + " , " + Progress + " , '" + BaseLineStart + "' , '" + BaseLineFinished + "' , '" + PlanStart + "' , '" + PlanFinished + "' , '" + ActualStart + "' , '" + ActualFinished + "' , " + OriginalDuration + " , " + SourceOfItem + " , " + ManPower + " , '" + LastTransmitalNumber + "' , '" + LastTransmitalDate + "' , '" + LastRevisionNumber + "' , " + LastStatus + " , " + LastAction + " )";
                        var ModifiedQuery = Query.Replace(", '' ", ", NULL").Replace(",  ,", ", NULL ,").Replace(",  ", ", NULL"); //Most Important Part Of Insert Code - This Part Modifies Code And Put Null In Places Where Data Is Empty
                        SqlCommand Command = new SqlCommand(ModifiedQuery, SQLConnection);
                        Command.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        ErroredTasks.Add(Index, Convert.ToString(DataTable.Rows[Index][4]));
                    }
                }

                else { ExistedTasks.Add(Index, Convert.ToString(DataTable.Rows[Index][4])); }

            }

            var NotInserted = new { ExistedTasks, ErroredTasks };

            SQLConnection.Close();

            Stream.Close();

            return Json(NotInserted);
        }


        public async Task<ActionResult> ExportExcel()
        {
            var Result = await ProjectsDocumentsFilter(2,
                    HttpContext.Session.GetInt32("SearchType") != null ? HttpContext.Session.GetInt32("SearchType") : null,
                    HttpContext.Session.GetInt32("ProjectID") != null ? HttpContext.Session.GetInt32("ProjectID") : null,
                    HttpContext.Session.GetString("ProjectCode") != null ? HttpContext.Session.GetString("ProjectCode") : null,
                    HttpContext.Session.GetString("CompanyName") != null ? HttpContext.Session.GetString("CompanyName") : null,
                    HttpContext.Session.GetString("DocumentTitle") != null ? HttpContext.Session.GetString("DocumentTitle") : null,
                    HttpContext.Session.GetString("DocumentNumber") != null ? HttpContext.Session.GetString("DocumentNumber") : null,
                    HttpContext.Session.GetInt32("Dicipline") != null ? HttpContext.Session.GetInt32("Dicipline") : null,
                    HttpContext.Session.GetString("ResponsiblePerson") != null ? HttpContext.Session.GetString("ResponsiblePerson") : null,
                    HttpContext.Session.GetInt32("DocumentType") != null ? HttpContext.Session.GetInt32("DocumentType") : null,
                    HttpContext.Session.GetString("Description") != null ? HttpContext.Session.GetString("Description") : null,
                    HttpContext.Session.GetInt32("Progress") != null ? HttpContext.Session.GetInt32("Progress") : null,
                    HttpContext.Session.GetString("BaseLineStart") != null ? HttpContext.Session.GetString("BaseLineStart") : null,
                    HttpContext.Session.GetString("BaseLineFinished") != null ? HttpContext.Session.GetString("BaseLineFinished") : null,
                    HttpContext.Session.GetInt32("OriginalDuration") != null ? HttpContext.Session.GetInt32("OriginalDuration") : null,
                    HttpContext.Session.GetInt32("SourceOfItem") != null ? HttpContext.Session.GetInt32("SourceOfItem") : null,
                    HttpContext.Session.GetInt32("ManPower") != null ? HttpContext.Session.GetInt32("ManPower") : null,
                    HttpContext.Session.GetString("Critical") == "True" ? true : (bool?)null,
                    HttpContext.Session.GetString("TransmitalNumber") != null ? HttpContext.Session.GetString("TransmitalNumber") : null,
                    HttpContext.Session.GetString("TransmitalDate") != null ? HttpContext.Session.GetString("TransmitalDate") : null,
                    HttpContext.Session.GetString("Start-TransmitalDate") != null ? HttpContext.Session.GetString("Start-TransmitalDate") : null,
                    HttpContext.Session.GetString("End-TransmitalDate") != null ? HttpContext.Session.GetString("End-TransmitalDate") : null,
                    HttpContext.Session.GetString("CommentSheetNumber") != null ? HttpContext.Session.GetString("CommentSheetNumber") : null,
                    HttpContext.Session.GetString("CommentSheetDate") != null ? HttpContext.Session.GetString("CommentSheetDate") : null,
                    HttpContext.Session.GetString("ReplySheetNumber") != null ? HttpContext.Session.GetString("ReplySheetNumber") : null,
                    HttpContext.Session.GetString("ReplySheetDate") != null ? HttpContext.Session.GetString("ReplySheetDate") : null,
                    HttpContext.Session.GetString("RevisionNumber") != null ? HttpContext.Session.GetString("RevisionNumber") : null,
                    HttpContext.Session.GetInt32("Status") != null ? HttpContext.Session.GetInt32("Status") : null,
                    HttpContext.Session.GetInt32("Action") != null ? HttpContext.Session.GetInt32("Action") : null
                    );

            var List = Result.Value;


            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            if (!System.IO.Directory.Exists(_hostingEnvironment.WebRootPath + "\\Excel\\ProjectsTasks-ExportExcel")) { System.IO.Directory.CreateDirectory(_hostingEnvironment.WebRootPath + "\\Excel\\ProjectsTasks-ExportExcel"); } //Checks If Excel's Upload Folder Exists - If It Doesn't Then Create It

            // Create the file using the FileInfo object
            var FileName = "گزارشات مهندسی - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToLongDateTimeInt() + ".xlsx";
            var FileNameForReturn = "گزارشات مهندسی - " + new PersianDateTime(DateTime.Parse(DateTime.Now.ToString())).ToShortDateString() + ".xlsx"; //This Format Of File Is Used At The End For Returning The File 
            var OutPutDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Excel/ProjectsTasks-ExportExcel/");

            var ExcelFile = new FileInfo(OutPutDirectory + FileName);

            using (var Package = new ExcelPackage(ExcelFile))
            {
                //Save The ExcelFile In Excel Folder In WWWRoot Path=
                string Excels = Path.Combine(_hostingEnvironment.WebRootPath, "Excel");
                string FilePath = Path.Combine(Excels, "ProjectsTasks-ExportExcel");

                //ExcelFile.Create(); 
                //ExcelFile.CopyTo(FilePath);  //Copy ExportExcel File In Excel Folder In WWWRoot 

                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = Package.Workbook.Worksheets.Add("ProjectsTasks - " + DateTime.Now.ToShortDateString());
                worksheet.HeaderFooter.FirstFooter.LeftAlignedText = string.Format("Generated: {0}", DateTime.Now.ToShortDateString());
                worksheet.Row(1).Height = 15;
                worksheet.View.RightToLeft = true;

                // Start adding the header
                worksheet.Cells[1, 1].Value = "Document ID";
                worksheet.Cells[1, 2].Value = "Project ID";
                worksheet.Cells[1, 3].Value = "Project Name";
                worksheet.Cells[1, 4].Value = "Project Code";
                worksheet.Cells[1, 5].Value = "Company Name";
                worksheet.Cells[1, 6].Value = "Document Title";
                worksheet.Cells[1, 7].Value = "Document Number";
                worksheet.Cells[1, 8].Value = "Dicipline";
                worksheet.Cells[1, 9].Value = "Responsible Person";
                worksheet.Cells[1, 10].Value = "Document Type";
                worksheet.Cells[1, 11].Value = "Description";
                worksheet.Cells[1, 12].Value = "Weight Factor";
                worksheet.Cells[1, 13].Value = "Progress";
                worksheet.Cells[1, 14].Value = "Base Line Start";
                worksheet.Cells[1, 15].Value = "Base Line Finished";
                worksheet.Cells[1, 16].Value = "Plan Start";
                worksheet.Cells[1, 17].Value = "Plan Finished";
                worksheet.Cells[1, 18].Value = "Actual Start";
                worksheet.Cells[1, 19].Value = "Actual Finished";
                worksheet.Cells[1, 20].Value = "Original Duration";
                worksheet.Cells[1, 21].Value = "Source Of Item";
                worksheet.Cells[1, 22].Value = "Man Power";
                worksheet.Cells[1, 23].Value = "Critical";
                worksheet.Cells[1, 24].Value = "Last Transmital Number";
                worksheet.Cells[1, 25].Value = "Last Transmital Date";
                worksheet.Cells[1, 26].Value = "Last Revision Number";
                worksheet.Cells[1, 27].Value = "Last Status";
                worksheet.Cells[1, 28].Value = "Last Action";


                //// Add the second row of header data
                using (var range = worksheet.Cells[1, 1, 1, 28])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.WhiteSmoke);
                    range.Style.ShrinkToFit = false;
                    range.Style.WrapText = true;
                    range.Style.Locked = true;
                }

                int rowNumber = 2;

                foreach (var Item in List)
                {
                    worksheet.Cells[rowNumber, 1].Value = 1;
                    worksheet.Cells[rowNumber, 2].Value = Item.ProjectID;
                    worksheet.Cells[rowNumber, 3].Value = _projectAppService.GetProjectByID(Item.ProjectID).Result.Select(a => a.Title).FirstOrDefault();
                    worksheet.Cells[rowNumber, 4].Value = Item.ProjectCode;
                    worksheet.Cells[rowNumber, 5].Value = Item.CompanyName;
                    worksheet.Cells[rowNumber, 6].Value = Item.DocumentTitle;
                    worksheet.Cells[rowNumber, 7].Value = Item.DocumentNumber;
                    worksheet.Cells[rowNumber, 8].Value = Item.Dicipline;
                    worksheet.Cells[rowNumber, 9].Value = Item.ResponsiblePerson;
                    worksheet.Cells[rowNumber, 10].Value = Item.DocumentType;
                    worksheet.Cells[rowNumber, 11].Value = Item.Description;
                    worksheet.Cells[rowNumber, 12].Value = Item.WeightFactor;
                    worksheet.Cells[rowNumber, 13].Value = Item.Progress;
                    worksheet.Cells[rowNumber, 14].Value = Item.BaseLineStart;
                    worksheet.Cells[rowNumber, 15].Value = Item.BaseLineFinished;
                    worksheet.Cells[rowNumber, 16].Value = Item.PlanStart;
                    worksheet.Cells[rowNumber, 17].Value = Item.PlanFinished;
                    worksheet.Cells[rowNumber, 18].Value = Item.ActualStart;
                    worksheet.Cells[rowNumber, 19].Value = Item.ActualFinished;
                    worksheet.Cells[rowNumber, 20].Value = Item.OriginalDuration;
                    worksheet.Cells[rowNumber, 21].Value = Item.SourceOfItem;
                    worksheet.Cells[rowNumber, 22].Value = Item.ManPower;
                    worksheet.Cells[rowNumber, 23].Value = Item.Critical;
                    worksheet.Cells[rowNumber, 24].Value = Item.LastTransmitalNumber;
                    worksheet.Cells[rowNumber, 25].Value = Item.LastTransmitalDate;
                    worksheet.Cells[rowNumber, 26].Value = Item.LastRevisionNumber;
                    worksheet.Cells[rowNumber, 27].Value = Item.LastStatus;
                    worksheet.Cells[rowNumber, 28].Value = Item.LastAction;

                    rowNumber++;
                }

                Package.Workbook.Properties.Title = "ProjectsTasks - ExportExcel";
                Package.Workbook.Properties.Author = "Dapna.CO";
                Package.Workbook.Properties.Company = "MSVCO";

                var stream = new MemoryStream();
                Package.SaveAs(stream);
                stream.Position = 0;

                const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                stream.Position = 0;
                return File(stream, contentType, FileNameForReturn);
            }
        }

        public IActionResult Test()
        {
            return View();
        }
    }
}
