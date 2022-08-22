using Abp.Domain.Entities.Auditing;
using Dapna.MSVPortal.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Dapna.MSVPortal.ProjectsDocumentations
{
    public class ProjectsTasksRevisions : FullAuditedEntity<int>
    {
        //CreatorUserID & RevisionID Will Be Filled By Entity Framework - ProjectID Field Is Only For ViewModel

        //Identifiers
        public int TaskID { get; set; }
        public string RevisionNumber { get; set; } //This Field Will Be Filled After TransmitalNumber Enters

        public string TransmitalNumber { get; set; } //User Enters The Transmital Number Then The Fields Below Will Be Filled With Katibe's Database's Data
        public DateTime? TransmitalDate { get; set; }
        public string CommentSheetNumber { get; set; }
        public DateTime? CommentSheetDate { get; set; }
        public string ReplySheetNumber { get; set; }
        public DateTime? ReplySheetDate { get; set; }
        public ProjectsTasksStatusTypes? Status { get; set; }
        public ProjectsTasksActionTypes? Action { get; set; }
    }
}
