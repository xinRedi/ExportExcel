EXCEL动态导出说明文档
        
        //导出演示
        static void Main(string[] args)
        {
            //导出数据
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
            list.Add(new Dictionary<string, object>() { { "Name", "小王" }, { "Sex", "男" }, { "Age", "20" } });
            list.Add(new Dictionary<string, object>() { { "Name", "小美" }, { "Sex", "女" }, { "Age", "21" } });
            list.Add(new Dictionary<string, object>() { { "Name", "小康" }, { "Sex", "男" }, { "Age", "19" } });

            //获取文件字节
            byte[] entity = GetFileData("E:\\Test.xlsx");
            //构造函数赋值
            GenerateExcelReport.HandleReportConfig(entity);
            //输出Excel
            GenerateExcelReport.GenerateExcel(list);

            Console.ReadKey();
        }

        /// <summary>
        /// 将文件转换成byte[] 数组
        /// </summary>
        /// <param name="fileUrl">文件路径文件名称</param>
        /// <returns>byte[]</returns>
        public static byte[] GetFileData(string fileUrl)
        {
            FileStream fs = new FileStream(fileUrl, FileMode.Open, FileAccess.Read);
            try
            {
                byte[] buffur = new byte[fs.Length];
                fs.Read(buffur, 0, (int)fs.Length);

                return buffur;
            }
            catch (Exception ex)
            {
                //MessageBoxHelper.ShowPrompt(ex.Message);
                return null;
            }
            finally
            {
                if (fs != null)
                {
                    //关闭资源
                    fs.Close();
                }
            }
        }
