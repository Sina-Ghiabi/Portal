using Abp.Application.Services;
using Abp.Domain.Repositories;
using Dapna.MSVPortal.Enums;
using Dapna.MSVPortal.ProjectsDocumentations.Dto;
using Dapna.MSVPortal.Web.ViewModels;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.ProjectsDocumentations
{
    public interface IProjectsTasksRevisionsAppService : IAsyncCrudAppService<ProjectsTasksRevisionsDto, int>
    {
        Task<List<ProjectsTasksRevisions>> GetAllProjectsTaskRevisions();
        Task<List<ProjectsTasksRevisions>> GetSpecificProjectsTaskRevisions(int ProjectTaskID);
        Task<List<ProjectsTasksRevisions>> CheckRevisionsExistance(int TaskID, string TransmitalNumber);
        Task<List<ProjectsTasksRevisions>> GetSpecificProjectsTasksRevision(int ProjectsTasksRevisionID);
        Task<List<ProjectsTasksRevisions>> RevisionsFilterResult(string TransmitalNumber, DateTime? TransmitalDate , DateTime? StartTranmitalDate , DateTime? EndTransmitalDate , string CommentSheetNumber, DateTime? CommentSheetDate, string ReplySheetNumber, DateTime? ReplySheetDate, string RevisionNumber, ProjectsTasksStatusTypes? Status, ProjectsTasksActionTypes? Action);
        Task<List<ProjectsTasksRevisions>> LastRevisionsFilterResult(List<int> RevisionsIDs, string TransmitalNumber, DateTime? TransmitalDate , DateTime? StartTranmitalDate , DateTime? EndTransmitalDate , string CommentSheetNumber, DateTime? CommentSheetDate, string ReplySheetNumber, DateTime? ReplySheetDate, string RevisionNumber, ProjectsTasksStatusTypes? Status, ProjectsTasksActionTypes? Action);
        Task<bool> DeleteProjectsTaskRevisions(int TaskID);

        //Dashboard
        Task<List<ProjectsTasksRevisions>> GetAFCRevisions();
        Task<List<ProjectsTasksRevisions>> AllDocumentsFilter(List<int> TaskIDs , DateTime? StartDate, DateTime? EndDate);
        Task<List<ProjectsTasksRevisions>> DiciplineProduceProcess(List<int> TaskIDs = null , DateTime? StartDate = null);
        Task<List<ProjectsTasksRevisions>> AverageProjectsDocumentsProducedRevisions(List<int> TaskIDs, DateTime? StartDate = null , DateTime? EndDate = null);
        Task<List<ProjectsTasksRevisions>> AllRevisionsInPeriodOfTime(DateTime? StartDate = null, DateTime? EndDate = null);
    }

    public class ProjectsTasksRevisionsAppService : AsyncCrudAppService<ProjectsTasksRevisions, ProjectsTasksRevisionsDto, int>, IProjectsTasksRevisionsAppService
    {
        public ProjectsTasksRevisionsAppService(IRepository<ProjectsTasksRevisions, int> repository)
            : base(repository)
        {
        }
        public async Task<List<ProjectsTasksRevisions>> GetAllProjectsTaskRevisions()
        {
            return await Repository.GetAll().OrderBy(a=>a.TaskID).ToListAsync();
        }
        public async Task<List<ProjectsTasksRevisions>> CheckRevisionsExistance(int TaskID, string TransmitalNumber)
        {
            return await Repository.GetAll().Where(a => a.TaskID == TaskID && a.TransmitalNumber == TransmitalNumber && a.IsDeleted == false).ToListAsync();
        }
        public async Task<List<ProjectsTasksRevisions>> GetSpecificProjectsTaskRevisions(int ProjectTaskID)
        {
            return await Repository.GetAll().Where(a => a.TaskID == ProjectTaskID && a.IsDeleted == false).OrderBy(a => a.TransmitalDate).ToListAsync();
        }
        public async Task<List<ProjectsTasksRevisions>> GetSpecificProjectsTasksRevision(int ProjectsTasksRevisionID)
        {
            return await Repository.GetAll().Where(a => a.Id == ProjectsTasksRevisionID).OrderBy(a => a.Id).ToListAsync();
        }
        public async Task<List<ProjectsTasksRevisions>> RevisionsFilterResult(string TransmitalNumber, DateTime? TransmitalDate, DateTime? StartTransmitalDate, DateTime? EndTransmitalDate, string CommentSheetNumber, DateTime? CommentSheetDate, string ReplySheetNumber, DateTime? ReplySheetDate, string RevisionNumber, ProjectsTasksStatusTypes? Status, ProjectsTasksActionTypes? Action)
        {
            var Query = Repository.GetAll();

            if (TransmitalNumber != null) { Query = Query.Where(a => a.TransmitalNumber.Contains(TransmitalNumber)); }

            if (TransmitalDate != null || StartTransmitalDate != null || EndTransmitalDate != null)
            {
                if (TransmitalDate != null) { Query = Query.Where(a => a.TransmitalDate == TransmitalDate); }

                if (StartTransmitalDate != null && EndTransmitalDate != null) { Query = Query.Where(a => a.TransmitalDate >= StartTransmitalDate && a.TransmitalDate <= EndTransmitalDate); }

                else if (StartTransmitalDate != null && EndTransmitalDate == null) { Query = Query.Where(a => a.TransmitalDate >= StartTransmitalDate); }

                else if (StartTransmitalDate == null && EndTransmitalDate != null) { Query = Query.Where(a => a.TransmitalDate <= EndTransmitalDate); }
            }

            if (CommentSheetNumber != null) { Query = Query.Where(a => a.CommentSheetNumber.Contains(CommentSheetNumber)); }

            if (CommentSheetDate.HasValue) { Query = Query.Where(a => a.CommentSheetDate == CommentSheetDate); }

            if (ReplySheetNumber != null) { Query = Query.Where(a => a.ReplySheetNumber.Contains(ReplySheetNumber)); }

            if (ReplySheetDate != null) { Query = Query.Where(a => a.ReplySheetDate == ReplySheetDate); }

            if (RevisionNumber != null) { Query = Query.Where(a => a.RevisionNumber == RevisionNumber); }

            if (Status.HasValue) { Query = Query.Where(a => a.Status == Status); }

            if (Action.HasValue) { Query = Query.Where(a => a.Action == Action); }

            return await Query.OrderBy(a=>a.TaskID).ToListAsync();
        }
        public async Task<List<ProjectsTasksRevisions>> LastRevisionsFilterResult(List<int> RevisionsIDs, string TransmitalNumber, DateTime? TransmitalDate, DateTime? StartTransmitalDate, DateTime? EndTransmitalDate, string CommentSheetNumber, DateTime? CommentSheetDate, string ReplySheetNumber, DateTime? ReplySheetDate, string RevisionNumber, ProjectsTasksStatusTypes? Status, ProjectsTasksActionTypes? Action)
        {
            var Query = Repository.GetAll().Where(a => RevisionsIDs.Contains(a.Id));

            if (TransmitalNumber != null) { Query = Query.Where(a => a.TransmitalNumber.Contains(TransmitalNumber)); }

            if (TransmitalDate != null || StartTransmitalDate != null || EndTransmitalDate != null)
            {
                if (TransmitalDate != null) { Query = Query.Where(a => a.TransmitalDate == TransmitalDate); }

                if (StartTransmitalDate != null && EndTransmitalDate != null) { Query = Query.Where(a => a.TransmitalDate >= StartTransmitalDate && a.TransmitalDate <= EndTransmitalDate); }

                else if (StartTransmitalDate != null && EndTransmitalDate == null) { Query = Query.Where(a => a.TransmitalDate >= StartTransmitalDate); }

                else if (StartTransmitalDate == null && EndTransmitalDate != null) { Query = Query.Where(a => a.TransmitalDate <= EndTransmitalDate); }
            }

            if (CommentSheetNumber != null) { Query = Query.Where(a => a.CommentSheetNumber.Contains(CommentSheetNumber)); }

            if (CommentSheetDate.HasValue) { Query = Query.Where(a => a.CommentSheetDate == CommentSheetDate); }

            if (ReplySheetNumber != null) { Query = Query.Where(a => a.ReplySheetNumber.Contains(ReplySheetNumber)); }

            if (ReplySheetDate != null) { Query = Query.Where(a => a.ReplySheetDate == ReplySheetDate); }

            if (RevisionNumber != null) { Query = Query.Where(a => a.RevisionNumber == RevisionNumber); }

            if (Status.HasValue) { Query = Query.Where(a => a.Status == Status); }

            if (Action.HasValue) { Query = Query.Where(a => a.Action == Action); }

            return await Query.ToListAsync();
        }


        public async Task<bool> DeleteProjectsTaskRevisions(int TaskID)
        {
            var TasksRevisions = Repository.GetAll().Where(a => a.TaskID == TaskID);
            foreach (var Item in TasksRevisions)
            {
                await Repository.DeleteAsync(Item);
            }
            return true;
        }


        //Dashboard
        public async Task<List<ProjectsTasksRevisions>> GetAFCRevisions()
        {
            var Query = Repository.GetAll().Where(a => a.Status == (ProjectsTasksStatusTypes)1);

            return await Query.ToListAsync();
        }
        public async Task<List<ProjectsTasksRevisions>> AllDocumentsFilter(List<int> TaskIDs, DateTime? StartDate, DateTime? EndDate)
        {
            var Query = Repository.GetAll().Where(a => TaskIDs.Contains(a.TaskID));

            if (StartDate.HasValue && EndDate.HasValue) { Query = Query.Where(a => a.TransmitalDate >= StartDate && a.TransmitalDate <= EndDate); }

            return await Query.ToListAsync();
        }

        public async Task<List<ProjectsTasksRevisions>> DiciplineProduceProcess(List<int> TaskIDs = null, DateTime? StartDate = null)
        {
            var Query = Repository.GetAll().Where(a => TaskIDs.Contains(a.TaskID));

            if (StartDate.HasValue)
            {
                Query = Query.Where(a => a.TransmitalDate >= StartDate);
            }
            return await Query.ToListAsync();
        }

        public async Task<List<ProjectsTasksRevisions>> AverageProjectsDocumentsProducedRevisions(List<int> TaskIDs , DateTime? StartDate = null , DateTime? EndDate = null)
        {
            var Query = Repository.GetAll().Where(a => TaskIDs.Contains(a.TaskID));

            if (StartDate.HasValue)
            {
                Query = Query.Where(a => a.TransmitalDate >= StartDate);
            }

            if (EndDate.HasValue)
            {
                Query = Query.Where(a => a.TransmitalDate <= EndDate);
            }

            return await Query.ToListAsync();
        }

        public async Task<List<ProjectsTasksRevisions>> AllRevisionsInPeriodOfTime(DateTime? StartDate = null, DateTime? EndDate = null)
        {
            var Query = Repository.GetAll();

            if (StartDate.HasValue)
            {
                Query = Query.Where(a => a.TransmitalDate >= StartDate);
            }

            if (EndDate.HasValue)
            {
                Query = Query.Where(a => a.TransmitalDate <= EndDate);
            }

            return await Query.ToListAsync();
        }
    }
}
