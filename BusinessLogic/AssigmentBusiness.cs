using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataAccess;
using Entity;
using Entity.Results;
using Entity.Common;

namespace BusinessLogic
{
    public class AssigmentBusiness
    {
        public AssignmentResult InsertAssignment(StatementResult statement, string userCode)
        {
            return new AssigmentData().InsertAssignment(statement, userCode);            
        }

        public StatementResult GenerateAssigncSentences(string fileNameWithLocation, Provider provider, string sheet)
        {
            return new AssigmentData().GenerateAssigncSentences(fileNameWithLocation, provider, sheet);
        }

        public void DeleteAssignment(string userCode)
        {
            new AssigmentData().DeleteAssignment(userCode);
        }
    }
}
