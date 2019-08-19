using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace TableMaker
{
    public partial class RibbonTab
    {
        public static readonly Color Green = Color.FromArgb(0, 176, 80);
        public static readonly Color Blue = Color.FromArgb(79, 129, 189);
        public static readonly Color Cyan = Color.FromArgb(31, 217, 243);
        public static readonly Color Orange = Color.FromArgb(226, 107, 10);

        public const int _RowDescription = 2;
        public const int _RowType = 3;
        public const int _RowLenOrExp = 4;
        public const int _RowField = 5;


        public Excel.Application app;
        private void RibbonTab_Load(object sender, RibbonUIEventArgs e)
        {
            app = ThisAddIn.app;
            app.SheetChange += App_SheetChange;
        }

        private void App_SheetChange(object Sh, Range Target)
        {

            if (!IsLoad() || !IsTable(app.ActiveSheet)) return;
            if (Target.Value != null)
            {
                if (Target.Value is object[])
                {
                    dynamic[] rangeArray = Target.Value;
                    for (int ri = 0; ri < rangeArray.Length; ri++)
                        if (rangeArray[ri] == null)
                        {

                        }

                }
                if (Target.Value is object[,])
                {
                    dynamic[,] rangeArray = Target.Value;

                    for (int ri = 1; ri <= rangeArray.Rank; ri++)
                        for (int ci = 1; ci < rangeArray.GetLength(ri-1); ci++)
                        {
                            int a = rangeArray.GetLength(ri - 1);
                            if (rangeArray[ri,ci] == null)
                            {

                            }
                        }
                }
                if (Target.Value is string)
                {
                    Target.Font.Name = "Microsoft YaHei UI";
                }

            }

            if (Target.Row == 1)
            {
                Target.Font.Color = Color.Gray;
            }
            if (Target.Row == _RowDescription)
            {
                Target.Font.Color = Color.Green;
            }
            if (Target.Row == _RowType)
            {
                Target.Font.Color = Blue;
                Target.Font.Bold = true;

                string targetType = ((Range)app.Cells[_RowType, Target.Column]).Value;
                if (targetType == string.Empty || targetType == null) targetType = "string";
                targetType = targetType.Trim().ToLower();

                if (targetType == "int" || targetType == "int32" || targetType == "integer")
                {
                    ((Range)app.Columns[Target.Column]).NumberFormatLocal = "0";
                }
                else if (targetType == "float" || targetType == "single")
                {
                }
                else
                {
                    ((Range)app.Columns[Target.Column]).NumberFormatLocal = "@";
                }
            }
            if (Target.Row == _RowLenOrExp)
            {
                Target.Font.Color = Orange;
                Target.NumberFormatLocal = "@";
            }
            if (Target.Row == _RowField)
            {
                int LastColumn = app.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                //MessageBox.Show(LastColumn.ToString());

                if (Target.Text == "")
                {
                    Target.Font.ColorIndex = 0;
                    Target.Interior.ColorIndex = 0;
                }
                else
                {
                    Target.Font.Color = Color.White;
                    Target.Interior.Color = Blue;
                }
            }
            if (Target.Row > _RowField)
                DataRegion(Sh, Target);
        }

        private void DataRegion(object Sh, Range Target)
        {
            if (Target == null) return;
            string targetType = ((Range)app.Cells[_RowType, Target.Column]).Value;
            if (targetType == string.Empty || targetType == null) targetType = "string";
            targetType = targetType.Trim().ToLower();

            if (targetType == "int" || targetType == "int32" || targetType == "integer")
            {
                Target.Font.Color = Color.Purple;
            }
            else if (targetType == "float" || targetType == "single")
            {
                Target.Font.Color = Color.DarkOrchid;
            }
            else if (targetType == "vector2")
            {
                Target.Font.Color = Color.DarkCyan;
            }
            else if (targetType == "vector3")
            {
                Target.Font.Color = Color.DarkCyan;
            }
            else if (targetType == "vector4")
            {
                Target.Font.Color = Color.DarkCyan;
            }
            else if (targetType == "transform")
            {
                Target.Font.Color = Cyan;
            }
            else if (targetType == "color")
            {
                Target.Font.Color = Color.DeepPink;
            }
            else //targetType == string or more..
            {
                Target.Font.Color = Orange;
            }

        }


        private void ErrorCheckBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = app.ActiveSheet;
            Range r = sheet.Cells[1, 1];
            MessageBox.Show(r.Text);
        }

        private bool IsTable(Excel.Worksheet sheet)
        {
            if (sheet.Cells[1, 1].Value == "$Table")
                return true;
            else
                return false;
        }
        private bool IsLoad() => IsLoadCheck.Checked;

        private void NewTableBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = app.ActiveSheet;
            sheet.Cells[1, 1].Value = "$Table";
            sheet.Cells[1, 2].Value = "$Name:" + sheet.Name;
            sheet.Cells[2, 1].Value = "字段1";
            sheet.Cells[2, 2].Value = "字段2";
            sheet.Cells[3, 1].Value = "int";
            sheet.Cells[3, 2].Value = "string";
            sheet.Cells[4, 1].Value = "20";
            sheet.Cells[4, 2].Value = "4";
            sheet.Cells[5, 1].Value = "ID";
            sheet.Cells[5, 2].Value = "Data";
            sheet.Cells[6, 1].Value = "18";
            sheet.Cells[6, 2].Value = "A String";
        }

        private void NewExmTableBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = app.Sheets.Add();
            //Excel.Worksheet sheet = app.ActiveSheet;
            sheet.Cells[1, 1].Value = "$Table";
            sheet.Cells[1, 2].Value = "$Name:" + sheet.Name;

            sheet.Cells[2, 1].Value = "ID序号";
            sheet.Cells[2, 2].Value = "名字";
            sheet.Cells[2, 3].Value = "攻击力";
            sheet.Cells[2, 4].Value = "图标路径";
            sheet.Cells[2, 5].Value = "图标UI位置";
            sheet.Cells[2, 6].Value = "出生点";
            sheet.Cells[2, 7].Value = "角色变换";
            sheet.Cells[2, 8].Value = "皮肤染色";
            sheet.Cells[2, 9].Value = "没用过的4";

            sheet.Cells[3, 1].Value = "int";
            sheet.Cells[3, 2].Value = "string";
            sheet.Cells[3, 3].Value = "float";
            sheet.Cells[3, 4].Value = "string";
            sheet.Cells[3, 5].Value = "Vector2";
            sheet.Cells[3, 6].Value = "Vector3";
            sheet.Cells[3, 7].Value = "Transform";
            sheet.Cells[3, 8].Value = "Color";
            sheet.Cells[3, 9].Value = "Vector4";

            sheet.Cells[4, 1].Value = "0~5";
            sheet.Cells[4, 2].Value = "4";
            sheet.Cells[4, 3].Value = "120~400";
            sheet.Cells[4, 4].Value = "$E:.*\\...\\.prefab";
            sheet.Cells[4, 5].Value = "0,0~1920,1080";
            sheet.Cells[4, 6].Value = "";
            sheet.Cells[4, 7].Value = "";
            sheet.Cells[4, 8].Value = "";
            sheet.Cells[4, 9].Value = "";
            ((Range)sheet.Cells[4, 4]).AddComment("字符串如果使用正则表达式匹配则需要在前面加上$E:");
            ((Range)sheet.Cells[4, 2]).AddComment("正常可以写2就是最大长度2\n3~5就是最小长度3最大长度5");

            sheet.Cells[5, 1].Value = "ID";
            sheet.Cells[5, 2].Value = "Name";
            sheet.Cells[5, 3].Value = "Attack";
            sheet.Cells[5, 4].Value = "UIPath";
            sheet.Cells[5, 5].Value = "UIPosition";
            sheet.Cells[5, 6].Value = "SpawnPosition";
            sheet.Cells[5, 7].Value = "RoleTransform";
            sheet.Cells[5, 8].Value = "RoleColor";
            sheet.Cells[5, 9].Value = "emmm";

            sheet.Cells[6, 1].Value = "1";
            sheet.Cells[6, 2].Value = "小鸟";
            sheet.Cells[6, 3].Value = "120.5";
            sheet.Cells[6, 4].Value = "Asset\\001.prefab";
            sheet.Cells[6, 5].Value = "500,500";
            sheet.Cells[6, 6].Value = "0,0,0";
            sheet.Cells[6, 7].Value = "50,50,50;0,0,0;1,1,1";
            sheet.Cells[6, 8].Value = "+255,255,255";
            sheet.Cells[6, 9].Value = "0,0,0,0";
            ((Range)sheet.Cells[6, 8]).AddComment("0~255的Color，同时支持RGBA，例如+255,255,255,255");

            sheet.Cells[7, 1].Value = "2";
            sheet.Cells[7, 2].Value = "小猫";
            sheet.Cells[7, 3].Value = "125.55";
            sheet.Cells[7, 4].Value = "Asset\\002.prefab";
            sheet.Cells[7, 5].Value = "450,400";
            sheet.Cells[7, 6].Value = "0,30,0";
            sheet.Cells[7, 7].Value = "25,25,25;0,30,0;1,1,1";
            sheet.Cells[7, 8].Value = "1,1,1";
            sheet.Cells[7, 9].Value = "0,0,0,0";
            ((Range)sheet.Cells[7, 8]).AddComment("0~1的RGB值，同时支持RGBA，例如1,1,1,1");

            sheet.Cells[8, 1].Value = "3";
            sheet.Cells[8, 2].Value = "小3";
            sheet.Cells[8, 3].Value = "399.55";
            sheet.Cells[8, 4].Value = "Asset\\003.prefab";
            sheet.Cells[8, 5].Value = "30,10";
            sheet.Cells[8, 6].Value = "0,30,0";
            sheet.Cells[8, 7].Value = "80,25,56;0,4,35;1,1,1";
            sheet.Cells[8, 8].Value = "#FFFFFF";
            sheet.Cells[8, 9].Value = "0,0,0,0";
            ((Range)sheet.Cells[8, 8]).AddComment("16进制的颜色值，同事支持RGBA，例如#FFFFFFFF");
        }
    }
}
