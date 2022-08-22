using Dapna.MSVPortal.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.Web.ViewModels
{
    public class ProjectsTasksRevisionsViewModel
    {
        public int? CreatorUserID { get; set; }

        //Identifiers
        public string FormType { get; set; }
        public int? RevisionID { get; set; }
        public int? ProjectID { get; set; } 
        public string ProjectCode { get; set; }
        public int TaskID { get; set; }
        public string DocumentNumber { get; set; }
        public string DocumentTitle { get; set; }
        public string RevisionNumber { get; set; }//This Field Will Be Filled After TransmitalNumber Enters

        public string TransmitalNumber { get; set; }//User Enters The Transmital Number Then The Fields Below Will Be Filled With Katibe's Database's Data
        public string TransmitalDate { get; set; }
        public string CommentSheetNumber { get; set; }
        public string CommentSheetDate { get; set; }
        public string ReplySheetNumber { get; set; }
        public string ReplySheetDate { get; set; }
        public ProjectsTasksStatusTypes? Status { get; set; }
        public ProjectsTasksActionTypes? Action { get; set; }
    }
}
