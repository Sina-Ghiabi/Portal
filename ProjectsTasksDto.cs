using Abp.Application.Services.Dto;
using Abp.AutoMapper;
using Dapna.MSVPortal.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Dapna.MSVPortal.ProjectsDocumentations.Dto
{
    [AutoMap(typeof(ProjectsTasks))]
    public class ProjectsTasksDto : FullAuditedEntityDto<int>
    {
        //CreatorUserID & TaskID Will Be Filled By Entity Framework

        //ProjectsTask Fields
        public int ProjectID { get; set; }
        public string ProjectCode { get; set; }
        public string CompanyName { get; set; }
        public string DocumentTitle { get; set; }
        public string DocumentNumber { get; set; }
        public ProjectsTasksDiciplineTypes Dicipline { get; set; }
        public string ResponsiblePerson { get; set; }
        public ProjectsTasksDocumentTypes? DocumentType { get; set; }
        public string Description { get; set; }
        public Nullable<bool> Critical { get; set; }

        //ProjectsTaskSchedule Fields
        public int? WeightFactor { get; set; }
        public int? Progress { get; set; }
        public DateTime? BaseLineStart { get; set; }
        public DateTime? BaseLineFinished { get; set; }
        public DateTime? PlanStart { get; set; }
        public DateTime? PlanFinished { get; set; }
        public DateTime? ActualStart { get; set; }
        public DateTime? ActualFinished { get; set; }
        public int? OriginalDuration { get; set; }
        public ProjectsTasksSourceOfItemTypes? SourceOfItem { get; set; }
        public int? ManPower { get; set; }

        //Fields Below Will Be Filled From Last Revision 
        public string LastTransmitalNumber { get; set; }
        public DateTime? LastTransmitalDate { get; set; }
        public string LastRevisionNumber { get; set; }
        public ProjectsTasksStatusTypes? LastStatus { get; set; }
        public ProjectsTasksActionTypes? LastAction { get; set; }
    }
}
