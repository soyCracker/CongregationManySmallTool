using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmallTool.Lib.Models.DelegationTool;
using SmallTool.Lib.Utils;

namespace SmallTool.Lib.Services
{
    public class DtRecordService
    {
        private string blockDate = "";

        public DtRecordService()
        {

        }

        public byte[] Start(string formFile, string assignFile)
        {
            using FileStream formFileFs = new FileStream(formFile, FileMode.Open);
            using MemoryStream formFileMs = new MemoryStream();
            formFileFs.CopyTo(formFileMs);
            List<DelegationVM> vmList = ReadDelegation(formFileMs);

            using FileStream assignFileFs = new FileStream(assignFile, FileMode.Open);
            using MemoryStream assignFileMs = new MemoryStream();
            assignFileFs.CopyTo(assignFileMs);
            return Record(vmList, assignFileMs);
        }

        public byte[] Start(MemoryStream formFile, MemoryStream assignFile)
        {
            List<DelegationVM> vmList = ReadDelegation(formFile);
            return Record(vmList, assignFile);
        }

        private List<DelegationVM> ReadDelegation(MemoryStream formFile)
        {
            try
            {
                IWorkbook workbook;
                // 避免報異常EOF in header
                formFile.Position = 0;
                try
                {
                    workbook = new XSSFWorkbook(formFile);
                }
                catch (Exception)
                {
                    workbook = new HSSFWorkbook(formFile);
                }
                ISheet sheet = workbook.GetSheetAt(0);
                List<DelegationVM> vmList = new List<DelegationVM>();
                for (int j = 1; j <= 2; j++)
                {
                    for (int i = 4; i <= sheet.LastRowNum; i++)
                    {
                        SetDelegationList(sheet, i, vmList, j);
                    }
                }
                workbook.Close();
                return vmList;
            }
            catch (IOException)
            {
                throw;
            }
        }

        private void SetDelegationList(ISheet sheet, int rowNum, List<DelegationVM> vmList, int classInt)
        {
            DelegationVM vm = new DelegationVM();
            try
            {
                vm.Name = sheet.GetRow(rowNum).GetCell(2 + (classInt - 1) * 2).ToString().Trim();
                if (string.IsNullOrEmpty(vm.Name))
                {
                    return;
                }
                Console.WriteLine("委派人: " + vm.Name);
            }
            catch (Exception)
            {
                return;
            }
            try
            {
                vm.Assistant = sheet.GetRow(rowNum).GetCell(3 + (classInt - 1) * 2).ToString();
            }
            catch (Exception)
            {
                vm.Assistant = "";
            }
            vm.Header = StringUtil.GetChinesePrintAble(sheet.GetRow(rowNum).GetCell(1).ToString());
            vm.Class = classInt.ToString();
            try
            {
                vm.Date = sheet.GetRow(rowNum).GetCell(0).ToString();
                if (string.IsNullOrEmpty(vm.Date))
                {
                    vm.Date = blockDate;
                }
                else
                {
                    blockDate = vm.Date;
                }
            }
            catch (Exception)
            {
                vm.Date = blockDate;
            }
            vmList.Add(vm);
        }

        private byte[] Record(List<DelegationVM> vmList, MemoryStream assignFile)
        {
            try
            {
                IWorkbook workbook;
                assignFile.Position = 0;
                try
                {
                    workbook = new XSSFWorkbook(assignFile);
                }
                catch (Exception)
                {
                    workbook = new HSSFWorkbook(assignFile);
                }
                ISheet broSheet = workbook.GetSheetAt(0);
                ISheet sisSheet = workbook.GetSheetAt(1);
                ISheet assSheet = workbook.GetSheetAt(2);
                foreach (DelegationVM vm in vmList)
                {
                    if (!RecordBroSheet(broSheet, vm))
                    {
                        RecordSisSheet(sisSheet, assSheet, vm);
                    }
                }
                MemoryStream ms = new MemoryStream();
                workbook.Write(ms, false);
                workbook.Close();
                return ms.ToArray();
            }
            catch (Exception)
            {
                throw;
            }
        }

        private bool RecordBroSheet(ISheet broSheet, DelegationVM vm)
        {
            int typeCol = 99;
            if (vm.Header.Contains(DelegationTypeName.Reading))
            {
                typeCol = 0;
            }
            else if (vm.Header.Contains(DelegationTypeName.Talk))
            {
                typeCol = 2;
            }
            else
            {
                return false;
            }

            for (int i = 2; i <= broSheet.LastRowNum; i++)
            {
                var row = broSheet.GetRow(i);
                if (row.GetCell(typeCol) != null && row.GetCell(typeCol).ToString().Trim() == vm.Name)
                {
                    //順便整理一下string
                    row.GetCell(typeCol).SetCellValue(vm.Name);
                    if (row.GetCell(typeCol + 1) == null)
                    {
                        row.CreateCell(typeCol + 1);
                    }
                    row.GetCell(typeCol + 1).SetCellValue(TimeUtil.CovertDateToFileNameStr(vm.Date));
                    return true;
                }
            }
            return false;
        }

        private bool RecordSisSheet(ISheet sisSheet, ISheet assSheet, DelegationVM vm)
        {
            for (int i = 1; i <= sisSheet.LastRowNum; i++)
            {
                var row = sisSheet.GetRow(i);
                if (row.GetCell(0) != null && row.GetCell(0).ToString().Trim() == vm.Name)
                {
                    //順便整理一下string
                    row.GetCell(0).SetCellValue(vm.Name);
                    if (row.GetCell(1) == null)
                    {
                        row.CreateCell(1);
                    }
                    row.GetCell(1).SetCellValue(TimeUtil.CovertDateToFileNameStr(vm.Date));
                    string type = "X";
                    if (vm.Header.Contains(DelegationTypeName.InitialCall2))
                    {
                        type = "初";
                    }
                    else if (vm.Header.Contains(DelegationTypeName.RV2))
                    {
                        type = "續";
                    }
                    else if (vm.Header.Contains(DelegationTypeName.BibleStudy3))
                    {
                        type = "課";
                    }
                    if (row.GetCell(2) == null)
                    {
                        row.CreateCell(2);
                    }
                    row.GetCell(2).SetCellValue(type);
                    RecordAssSheet(assSheet, vm);
                    return true;
                }
            }
            return false;
        }

        private void RecordAssSheet(ISheet assSheet, DelegationVM vm)
        {
            for (int i = 1; i <= assSheet.LastRowNum; i++)
            {
                var row = assSheet.GetRow(i);
                if (row.GetCell(0) != null && row.GetCell(0).ToString().Trim() == vm.Assistant)
                {
                    //順便整理一下string
                    row.GetCell(0).SetCellValue(vm.Assistant);
                    if (row.GetCell(1) == null)
                    {
                        row.CreateCell(1);
                    }
                    row.GetCell(1).SetCellValue(TimeUtil.CovertDateToFileNameStr(vm.Date));
                    for (int j = 2; j <= 7; j++)
                    {
                        if (row.GetCell(j) == null)
                        {
                            row.CreateCell(j);
                            row.GetCell(j).SetCellValue(vm.Name);
                            return;
                        }
                        if (string.IsNullOrEmpty(row.GetCell(j).ToString().Trim()))
                        {
                            row.GetCell(j).SetCellValue(vm.Name);
                            return;
                        }
                        if (j == 7)
                        {
                            row.GetCell(2).SetCellValue(vm.Name);
                            for (int k = 3; k <= 7; k++)
                            {
                                row.GetCell(k).SetCellValue("");
                            }
                            return;
                        }
                    }
                }
            }
        }
    }
}
