using Abp.Application.Services;
using Abp.Domain.Repositories;
using Dapna.MSVPortal.Financial.Dto;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.Financial
{
    public interface IRemainCreditAppService : IAsyncCrudAppService<RemainCreditDto, int>
    {
        Task<List<RemainCredit>> GetUsersAllProjects(List<int> UsersProjects);
        Task<List<RemainCredit>> GetUsersFilteredProjects(List<int> UsersProjects, List<int> ProjectIDs, string PayTo, string FourthLevelCode , int SortType , int ShowAll );
        Task<List<RemainCredit>> GetTotalDebtAndCredit(List<int> UsersProjects);
        Task<List<RemainCredit>> GetFilteredTotalDebtAndCredit(List<int> UsersProjects , List<int?> ProjectIDs, string PayTo, string FourthLevelCode );
    }
    public class RemainCreditAppService : AsyncCrudAppService<RemainCredit, RemainCreditDto, int>, IRemainCreditAppService
    {
        public RemainCreditAppService (IRepository<RemainCredit, int> repository)
            : base(repository)
        {
        }

        public async Task<List<RemainCredit>> GetUsersAllProjects(List<int> UsersProjects)
        {
            return await Repository.GetAll().Where(a => UsersProjects.Contains(a.ProjectID)).ToListAsync();
        }
        public async Task<List<RemainCredit>> GetUsersFilteredProjects(List<int> UsersProjects, List<int> ProjectIDs, string PayTo, string FourthLevelCode , int SortType ,int ShowAll)
        {
            var Query = Repository.GetAll().Where(a => UsersProjects.Contains(a.ProjectID));

            if (ProjectIDs.Count() > 0) { Query = Query.Where(a => ProjectIDs.Contains(a.ProjectID)); }

            if (PayTo != null) { Query = Query.Where(a => a.PayTo.Contains(PayTo)); }

            if (FourthLevelCode != null) { Query = Query.Where(a => a.FourthLevelCode.Contains(FourthLevelCode)); }

            if (SortType != 0) 
            {
                if (SortType == 1) { Query = Query.OrderBy(a => a.RemainDebtAmount); }

                if (SortType == 2) { Query = Query.OrderByDescending(a => a.RemainDebtAmount); }

                if (SortType == 3) { Query = Query.OrderBy(a => a.RemainCreditAmount); }

                if (SortType == 4) { Query = Query.OrderByDescending(a => a.RemainCreditAmount); }
            }

            if (ShowAll == 1) { Query = Query.Where(a => a.RemainDebtAmount > 0 || a.RemainCreditAmount > 0); }

            return await Query.ToListAsync();
        }
        public async Task<List<RemainCredit>> GetTotalDebtAndCredit(List<int> UsersProjects)
        {
            return await Repository.GetAll().Where(a => UsersProjects.Contains(a.ProjectID)).OrderByDescending(a => a.Id).ToListAsync();
        }
        public async Task<List<RemainCredit>> GetFilteredTotalDebtAndCredit(List<int> UsersProjects, List<int?> ProjectIDs, string PayTo , string FourthLevelCode)
        {
            var Query = Repository.GetAll().Where(a => UsersProjects.Contains(a.ProjectID));

            if (ProjectIDs.Count > 0) { Query = Query.Where(a => ProjectIDs.Contains(a.ProjectID)); }

            if (PayTo != null) { Query = Query.Where(a => a.PayTo.Contains(PayTo)); }

            if (FourthLevelCode != null) { Query = Query.Where(a => a.FourthLevelCode.Contains(FourthLevelCode)); }

            return await Query.OrderByDescending(a => a.Id).ToListAsync();
        }
    }
}
