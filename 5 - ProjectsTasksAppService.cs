using Abp.Application.Services;
using Abp.Domain.Repositories;
using Dapna.MSVPortal.Enums;
using Dapna.MSVPortal.ProjectsDocumentations.Dto;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.ProjectsDocumentations
{
    public interface IProjectsTasksAppService : IAsyncCrudAppService<ProjectsTasksDto, int>
    {
        Task<List<ProjectsTasks>> GetAllProjectsTasks();
        Task<List<ProjectsTasks>> GetSpecificProjectsTask(int ProjectTaskID);
        Task<List<ProjectsTasks>> CheckTaskExistence(int ProjectID ,string DocumentTitle);
        Task<List<ProjectsTasks>> TasksFilterResult(int? ProjectID, string ProjectCode, string CompanyName, string DocumentTitle, string DocumentNumber, ProjectsTasksDiciplineTypes? Dicipline, string ResponsiblePerson, ProjectsTasksDocumentTypes? DocumentType, string Description, int? Progress, DateTime? BaseLineStart, DateTime? BaseLineFinished, int? OriginalDuration, ProjectsTasksSourceOfItemTypes? SourceOfItem, int? ManPower, bool? Critical);
        Task<List<ProjectsTasks>> ListedTasksFilterResult(List<int> RevisionsProjectIDs, int? ProjectID , string ProjectCode, string CompanyName, string DocumentTitle, string DocumentNumber, ProjectsTasksDiciplineTypes? Dicipline, string ResponsiblePerson, ProjectsTasksDocumentTypes? DocumentType, string Description, int? Progress, DateTime? BaseLineStart, DateTime? BaseLineFinished, int? OriginalDuration, ProjectsTasksSourceOfItemTypes? SourceOfItem, int? ManPower, bool? Critical);

        //Dashboard Charts
        Task<List<ProjectsTasks>> AFCDocumentsFilter(int? ProjectID , DateTime? StartDate , DateTime? EndDate);
        Task<List<ProjectsTasks>> AllProjectsDiciplineStatus(int ProjectID , DateTime StartDate , DateTime EndDate);
        Task<List<ProjectsTasks>> AverageAFCDocuments(int? ProjectID, DateTime? StartDate, DateTime? EndDate);
        Task<List<ProjectsTasks>> AverageToGetAFC(int? ProjectID, DateTime? StartDate, DateTime? EndDate);
        Task<List<ProjectsTasks>> GetDocumentsByProjectID(int ProjectID);
        Task<List<ProjectsTasks>> DiciplineProduceProcess(int? ProjectID = null , ProjectsTasksDiciplineTypes? Dicipline = null);
        Task<List<ProjectsTasks>> GetDocumentsWithAFCRevisions(List<int> AFCRevisionsTaskIDs);
        Task<List<ProjectsTasks>> SelectedAFCDocument(int? ProjectID, DateTime? StartDate, DateTime? EndDate);
    }

    public class ProjectsTasksAppService : AsyncCrudAppService<ProjectsTasks, ProjectsTasksDto, int>, IProjectsTasksAppService
    {
        public ProjectsTasksAppService(IRepository<ProjectsTasks, int> repository)
            : base(repository)
        {
        }
        public async Task<List<ProjectsTasks>> GetAllProjectsTasks()
        {
            return await Repository.GetAll().OrderBy(a=>a.Id).ToListAsync();
        }
        public async Task<List<ProjectsTasks>> GetSpecificProjectsTask(int ProjectTaskID)
        {
            return await Repository.GetAll().Where(a=>a.Id == ProjectTaskID && a.IsDeleted == false).OrderByDescending(a => a.Id).ToListAsync();
        }
        public async Task<List<ProjectsTasks>> CheckTaskExistence(int ProjectID , string DocumentTitle)
        {
            return await Repository.GetAll().Where(a => a.ProjectID == ProjectID && a.DocumentTitle == DocumentTitle  && a.IsDeleted == false).OrderByDescending(a => a.Id).ToListAsync();
        }
        public async Task<List<ProjectsTasks>> TasksFilterResult (int? ProjectID, string ProjectCode, string CompanyName, string DocumentTitle, string DocumentNumber, ProjectsTasksDiciplineTypes? Dicipline, string ResponsiblePerson, ProjectsTasksDocumentTypes? DocumentType, string Description, int? Progress, DateTime? BaseLineStart, DateTime? BaseLineFinished, int? OriginalDuration, ProjectsTasksSourceOfItemTypes? SourceOfItem, int? ManPower, bool? Critical)
        {
            var Query = Repository.GetAll();

            if (ProjectID.HasValue) { Query = Query.Where(a => a.ProjectID == ProjectID); }

            if (ProjectCode != null) { Query = Query.Where(a => a.ProjectCode.Contains(ProjectCode)); }

            if (CompanyName != null) { Query = Query.Where(a => a.CompanyName.Contains(CompanyName)); }

            if (DocumentTitle != null) { Query = Query.Where(a => a.DocumentTitle.Contains(DocumentTitle)); }

            if (DocumentNumber != null) { Query = Query.Where(a => a.DocumentNumber.Contains(DocumentNumber)); }

            if (Dicipline.HasValue) { Query = Query.Where(a => a.Dicipline == Dicipline); }

            if (ResponsiblePerson != null) { Query = Query.Where(a => a.ResponsiblePerson.Contains(ResponsiblePerson)); }

            if (DocumentType.HasValue) { Query = Query.Where(a => a.DocumentType == DocumentType); }

            if (Description != null) { Query = Query.Where(a => a.Description.Contains(Description)); }

            if (Progress.HasValue) { Query = Query.Where(a => a.Progress == Progress); }

            if (BaseLineStart.HasValue) { Query = Query.Where(a => a.BaseLineStart == BaseLineStart); }

            if (BaseLineFinished.HasValue) { Query = Query.Where(a => a.BaseLineFinished == BaseLineFinished); }

            if (OriginalDuration.HasValue) { Query = Query.Where(a => a.OriginalDuration == OriginalDuration); }

            if (SourceOfItem.HasValue) { Query = Query.Where(a => a.SourceOfItem == SourceOfItem); }

            if (ManPower.HasValue) { Query = Query.Where(a => a.ManPower == ManPower); }

            if (Critical.HasValue) { Query = Query.Where(a => a.Critical == Critical); }

            return await Query.OrderBy(a=>a.Id).ToListAsync();
        }


        public async Task<List<ProjectsTasks>> ListedTasksFilterResult(List<int> RevisionsProjectIDs, int? ProjectID , string ProjectCode, string CompanyName, string DocumentTitle, string DocumentNumber, ProjectsTasksDiciplineTypes? Dicipline, string ResponsiblePerson, ProjectsTasksDocumentTypes? DocumentType, string Description, int? Progress, DateTime? BaseLineStart, DateTime? BaseLineFinished, int? OriginalDuration, ProjectsTasksSourceOfItemTypes? SourceOfItem, int? ManPower, bool? Critical)
        {
            var Query = Repository.GetAll().Where(a => RevisionsProjectIDs.Contains(a.Id));

            if (ProjectID != null) { Query = Query.Where(a => a.ProjectID == ProjectID); }

            if (ProjectCode != null) { Query = Query.Where(a => a.ProjectCode.Contains(ProjectCode)); }

            if (CompanyName != null) { Query = Query.Where(a => a.CompanyName.Contains(CompanyName)); }

            if (DocumentTitle != null) { Query = Query.Where(a => a.DocumentTitle.Contains(DocumentTitle)); }

            if (DocumentNumber != null) { Query = Query.Where(a => a.DocumentNumber.Contains(DocumentNumber)); }

            if (Dicipline.HasValue) { Query = Query.Where(a => a.Dicipline == Dicipline); }

            if (ResponsiblePerson != null) { Query = Query.Where(a => a.ResponsiblePerson.Contains(ResponsiblePerson)); }

            if (DocumentType.HasValue) { Query = Query.Where(a => a.DocumentType == DocumentType); }

            if (Description != null) { Query = Query.Where(a => a.Description.Contains(Description)); }

            if (Progress.HasValue) { Query = Query.Where(a => a.Progress == Progress); }

            if (BaseLineStart.HasValue) { Query = Query.Where(a => a.BaseLineStart == BaseLineStart); }

            if (BaseLineFinished.HasValue) { Query = Query.Where(a => a.BaseLineFinished == BaseLineFinished); }

            if (OriginalDuration.HasValue) { Query = Query.Where(a => a.OriginalDuration == OriginalDuration); }

            if (SourceOfItem.HasValue) { Query = Query.Where(a => a.SourceOfItem == SourceOfItem); }

            if (ManPower.HasValue) { Query = Query.Where(a => a.ManPower == ManPower); }

            if (Critical.HasValue) { Query = Query.Where(a => a.Critical == Critical); }

            return await Query.ToListAsync();
        }


        //Dashboard Charts
        public async Task<List<ProjectsTasks>> GetDocumentsWithAFCRevisions(List<int> AFCRevisionsTaskIDs)
        {
            var Query = Repository.GetAll().Where(a => AFCRevisionsTaskIDs.Contains(a.Id));

            return await Query.OrderByDescending(a => a.Id).ToListAsync();
        }

        public async Task<List<ProjectsTasks>> AFCDocumentsFilter(int? ProjectID , DateTime? StartTime , DateTime? EndDate)
        {
            var Query = Repository.GetAll().Where(a => a.LastStatus == (ProjectsTasksStatusTypes)1 && a.IsDeleted == false);

            if (ProjectID.HasValue) { Query = Query.Where(a => a.ProjectID == ProjectID); }

            if (StartTime.HasValue && EndDate.HasValue ) { Query = Query.Where(a => a.LastTransmitalDate >= StartTime && a.LastTransmitalDate <= EndDate); }

            return await Query.OrderByDescending(a => a.Id).ToListAsync();
        }

        public async Task<List<ProjectsTasks>> AllProjectsDiciplineStatus(int ProjectID , DateTime StartDate , DateTime EndDate)
        {
            return await Repository.GetAll().Where(a => a.ProjectID == ProjectID && a.LastTransmitalDate >= StartDate && a.LastTransmitalDate <= EndDate && a.IsDeleted == false).OrderByDescending(a => a.Id).ToListAsync();
        }

        public async Task<List<ProjectsTasks>> AverageAFCDocuments(int? ProjectID , DateTime? StartDate , DateTime? EndDate)
        {
            var Query = Repository.GetAll();

            if (ProjectID.HasValue) { Query = Query.Where(a => a.ProjectID == ProjectID); };

            if (StartDate.HasValue && EndDate.HasValue) { Query = Query.Where(a => a.LastTransmitalDate >= StartDate && a.LastTransmitalDate <= EndDate); };

            return await Query.OrderByDescending(a => a.Id).ToListAsync();
        }
        public async Task<List<ProjectsTasks>> AverageToGetAFC(int? ProjectID, DateTime? StartDate, DateTime? EndDate)
        {
            var Query = Repository.GetAll().Where(a => a.LastStatus == (ProjectsTasksStatusTypes)1);

            if (ProjectID.HasValue) { Query = Query.Where(a => a.ProjectID == ProjectID); };

            if (StartDate.HasValue && EndDate.HasValue) { Query = Query.Where(a => a.LastTransmitalDate >= StartDate && a.LastTransmitalDate <= EndDate); };

            return await Query.OrderBy(a => a.LastTransmitalDate).ToListAsync();
        }
        public async Task<List<ProjectsTasks>> GetDocumentsByProjectID(int ProjectID)
        {
            return await Repository.GetAll().Where(a => a.ProjectID == ProjectID && a.IsDeleted == false).OrderByDescending(a => a.Id).ToListAsync();
        }

        public async Task<List<ProjectsTasks>> DiciplineProduceProcess(int? ProjectID = null , ProjectsTasksDiciplineTypes? Dicipline = null)
        {
            var Query = Repository.GetAll();

            if (ProjectID.HasValue)
            {
                Query = Query.Where(a => a.ProjectID == ProjectID);
            }

            if (Dicipline.HasValue)
            {
                Query = Query.Where(a => a.Dicipline == Dicipline);
            }

            return Query.ToList();
        }

        public async Task<List<ProjectsTasks>> SelectedAFCDocument(int? ProjectID, DateTime? StartDate, DateTime? EndDate)
        {
            var Query = Repository.GetAll();

            if (ProjectID.HasValue)
            {
                Query = Query.Where(a => a.ProjectID == ProjectID);
            }

            if (StartDate.HasValue && EndDate.HasValue )
            {
                Query = Query.Where(a => a.LastTransmitalDate >= StartDate && a.LastTransmitalDate <= EndDate);
            }

            return Query.ToList();
        }
    }
}
