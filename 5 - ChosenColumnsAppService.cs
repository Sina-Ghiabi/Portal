using Abp.Application.Services;
using Abp.Domain.Repositories;
using Dapna.MSVPortal.Projects.Dto;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dapna.MSVPortal.Projects
{
    public interface IChosenColumnsAppService : IAsyncCrudAppService<ChosenColumnsDto, int>
    {
        Task<List<ChosenColumns>> GetChosenColumns(long UserID);
        //Task<List<ChosenColumns>> GetChosenColumns();
    }


    public class ChosenColumnsAppService : AsyncCrudAppService<ChosenColumns, ChosenColumnsDto, int>, IChosenColumnsAppService
    {
        public ChosenColumnsAppService(IRepository<ChosenColumns, int> repository)
            : base(repository)
        {
        }

        public async Task<List<ChosenColumns>> GetChosenColumns(long UserID)
        {
            return await Repository.GetAll().Where(a => a.UserID == UserID).OrderByDescending(a => a.Id).ToListAsync();
            //return await Repository.GetAll().OrderByDescending(a => a.Id).ToListAsync();
        }

    }

}
