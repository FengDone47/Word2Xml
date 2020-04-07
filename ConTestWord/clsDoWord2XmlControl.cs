using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ConTestWord
{
    class clsDoWord2XmlControl
    {
        public string doWord2Txt(string docPath, string draPath, bool bCopyDra = true)
        {
            //string docPath = @"D:\每日工作处理文件\20200323\复审无效样例\无效\DOC\W36229_109752_6W109752.doc";
            //string draPath = @"D:\每日工作处理文件\20200323\复审无效样例\无效\DOC";
            string txtPath = draPath + "\\" + Path.GetFileNameWithoutExtension(docPath) + ".txt";
            Document document = new Document(docPath);
            int index = 0;
            string txtCotent = document.GetText();
            StreamWriter sw = new StreamWriter(txtPath, false, Encoding.UTF8);
            int iTable = 0;
            //document.SaveToFile("d:\\Targetxml.xml", FileFormat.WordXml);
            //Get Each Section of Document  
            foreach (Section section in document.Sections)
            {
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        #region Paragraph
                        Paragraph paragraph = obj as Paragraph;
                        StringBuilder sbPara = new StringBuilder();
                        //WORD默认生成的下拉小序号
                        if (!String.IsNullOrEmpty(paragraph.ListText))
                            sbPara.Append(paragraph.ListText);
                        foreach (DocumentObject docObject in paragraph.ChildObjects)
                        {
                            //If Type of Document Object is Picture, Extract.  
                            if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                            {
                                DocPicture pic = docObject as DocPicture;
                                String imgName = draPath + String.Format("\\Image -{0}.png", index);

                                //Save Image
                                sbPara.Append("<image>" + index + ".png" + "</image>");
                                if (bCopyDra)
                                    pic.Image.Save(imgName, System.Drawing.Imaging.ImageFormat.Png);
                                index++;
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange txt = docObject as TextRange;
                                if (txt.CharacterFormat.SubSuperScript == SubSuperScript.SubScript)
                                    sbPara.Append("<sub>" + txt.Text + "</sub>");
                                else if (txt.CharacterFormat.SubSuperScript == SubSuperScript.SuperScript)
                                    sbPara.Append("<sup>" + txt.Text + "</sup>");
                                else if (txt.CharacterFormat.SubSuperScript == SubSuperScript.BaseLine)
                                    sbPara.Append(txt.Text);
                                else
                                    sbPara.Append(txt.Text);
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.BookmarkStart)
                            {
                                sbPara.Append("<bookmark>");
                                continue;
                                #region 注释BOOKMARK
                                /*
                                BookmarkStart mark = docObject as BookmarkStart;
                                BookmarksNavigator navigator = new BookmarksNavigator(document);
                                navigator.MoveToBookmark(mark.Name);
                                TextBodyPart textBodyPart = navigator.GetBookmarkContent();
                                string text = null;
                                foreach (var item in textBodyPart.BodyItems)
                                {
                                    if (item is Paragraph)
                                    {
                                        text += (item as Paragraph).Text;
                                        foreach (var childObject in (item as Paragraph).ChildObjects)
                                        {
                                            if (childObject is TextRange)
                                            {
                                                text += (childObject as TextRange).Text;
                                            }
                                        }
                                    }
                                }
                                sbPara.Append("<bookmarkStart>" + text + "</bookmarkStart>");
                                */
                                #endregion
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.BookmarkEnd)
                            {
                                sbPara.Append("</bookmark>");
                                continue;
                                #region 注释BOOKMARK
                                /*
                                BookmarkEnd mark = docObject as BookmarkEnd;
                                BookmarksNavigator navigator = new BookmarksNavigator(document);
                                navigator.MoveToBookmark(mark.Name);
                                TextBodyPart textBodyPart = navigator.GetBookmarkContent();
                                string text = null;
                                foreach (var item in textBodyPart.BodyItems)
                                {
                                    if (item is Paragraph)
                                    {
                                        text += (item as Paragraph).Text;
                                        foreach (var childObject in (item as Paragraph).ChildObjects)
                                        {
                                            if (childObject is TextRange)
                                            {
                                                text += (childObject as TextRange).Text;
                                            }
                                        }
                                    }
                                }
                                sbPara.Append("<bookmarkEnd>" + text + "</bookmarkEnd>");
                                */
                                #endregion
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.TextFormField)
                            {
                                continue;
                                //TextFormField textForm = docObject as TextFormField;
                                //sbPara.Append("<textForm>" + textForm.Name + "</textForm>");
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.FieldMark)
                            {
                                //FieldMark fieldMark = docObject as FieldMark;
                                //sbPara.Append("<FieldMark>" + fieldMark.StyleName + "</FieldMark>");
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.CheckBox)
                            {
                                CheckBoxFormField checkBox = docObject as CheckBoxFormField;
                                sbPara.Append("<CheckBox>" + checkBox.Checked.ToString() + "</CheckBox>");
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.Break)
                            {
                                Break bk = docObject as Break;
                                if (sbPara.Length > 0 && (bk.BreakType == BreakType.PageBreak || bk.BreakType == BreakType.LineBreak))//无效数据中总是不折行
                                    sbPara.Append("\r\n");
                                sbPara.Append("<Break>" + bk.BreakType + "</Break>");
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.OleObject)
                            {
                                DocOleObject ole = docObject as DocOleObject;
                                sbPara.Append("<OleObject>" + ole.ShapeType + "</OleObject>");
                            }
                            else if (docObject.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange tr = docObject as TextRange;
                                sbPara.Append(tr.Text);
                            }
                            else
                            {

                            }
                        }
                        sw.WriteLine(sbPara.ToString());
                        #endregion
                    }
                    else if (obj.DocumentObjectType == DocumentObjectType.Table)
                    {
                        iTable++;
                        if (iTable > 2)//table需要转PIC，不再做内容了
                        {
                            sw.WriteLine("<image>" + index++ + ".png" + "</image>");
                            continue;
                        }
                        StringBuilder sbTable = new StringBuilder();
                        sbTable.Append("<table>");
                        Table table = obj as Table;
                        foreach (TableRow row in table.Rows)
                        {
                            sbTable.Append("<row>");
                            foreach (TableCell cell in row.Cells)
                            {
                                sbTable.Append("<col>");
                                foreach (DocumentObject item in cell.ChildObjects)
                                {
                                    if (item.DocumentObjectType == DocumentObjectType.Table)
                                    {
                                        continue;
                                    }
                                    Paragraph itemPara = item as Paragraph;
                                    foreach (DocumentObject child in itemPara.ChildObjects)
                                    {
                                        if (child.DocumentObjectType == DocumentObjectType.TextRange)
                                        {
                                            TextRange tr = child as TextRange;
                                            sbTable.Append(tr.Text);
                                        }
                                        else if (child.DocumentObjectType == DocumentObjectType.BookmarkStart)
                                        {
                                            sbTable.Append("<bookmark>");
                                        }
                                        else if (child.DocumentObjectType == DocumentObjectType.BookmarkEnd)
                                        {
                                            sbTable.Append("</bookmark>");
                                        }
                                        else if (child.DocumentObjectType == DocumentObjectType.FieldMark)
                                        {
                                            sbTable.Append("<FieldMark>");
                                        }
                                    }
                                }
                                sbTable.Append("</col>");
                            }
                            sbTable.Append("</row>");
                        }
                        sbTable.Append("</table>");
                        sw.WriteLine(sbTable.ToString());
                        sw.Flush();
                    }

                }
            }
            sw.Flush();
            sw.Close();
            sw.Dispose();
            return txtPath;
        }

        public void doTxt2Xml(string txtPath, string xmlPath)
        {
            StreamWriter sw = new StreamWriter(xmlPath, false, new UTF8Encoding(false));
            sw.WriteLine("<?xml version='1.0' encoding='UTF-8'?>");
            sw.WriteLine("<!DOCTYPE cn-appeal-decision SYSTEM 'cn-appeal-decision-v1-0.dtd'>");
            sw.WriteLine("<cn-appeal-decision>");
            sw.WriteLine("\t<cn-case-info>");
            sw.WriteLine("\t</cn-case-info>");

            using (StreamReader sr = new StreamReader(txtPath, Encoding.UTF8))
            {
                string content = "";
                bool bBegin = false; bool bEnd = false;
                int iNum = 1;
                while ((content = sr.ReadLine()) != null)
                {
                    content = Regex.Replace(content, @"[\x00-\x1f]", "");
                    content = content.Replace("<Break>PageBreak</Break>", "");
                    content = content.Replace("<Break>LineBreak</Break>", "");
                    content = content.Replace("<bookmark>", "");
                    content = content.Replace("</bookmark>", "");
                    content = content.Replace("<OleObject>MinValue</OleObject>", "");

                    if (String.IsNullOrEmpty(content))
                        continue;

                    if (Regex.IsMatch(content, "^[ 一、．.]{0,}案由[ ]{0,}$"))
                    {
                        bBegin = true;
                        sw.WriteLine("\t<cn-decision-detail>");
                        sw.WriteLine("\t\t<!--案由-->");
                        sw.WriteLine("\t\t<cn-brief-history>");
                        sw.WriteLine("\t\t\t<heading id=\"h01\"><b>一、案由</b></heading>");
                        continue;
                    }
                    else if (Regex.IsMatch(content, "^[ 二、．.]{0,}决定[的]?理由[ ]{0,}$"))
                    {
                        sw.WriteLine("\t\t</cn-brief-history>");
                        sw.WriteLine("\t\t<!--决定的理由-->");
                        sw.WriteLine("\t\t<cn-reasoning>");
                        sw.WriteLine("\t\t\t<heading id=\"h02\"><b>二、决定的理由</b></heading>");
                        continue;
                    }
                    else if (Regex.IsMatch(content, "^[ 三四、．.]{0,}决定[ ]{0,}$"))
                    {
                        sw.WriteLine("\t\t</cn-reasoning>");
                        sw.WriteLine("\t\t<!--决定-->");
                        sw.WriteLine("\t\t<cn-holding>");
                        sw.WriteLine("\t\t\t<heading id=\"h03\"><b>三、决定</b></heading>");
                        continue;
                    }
                    if (!bBegin)
                        continue;
                    if (bBegin && content.IndexOf("合议组组长：") == 0)
                        bEnd = true;
                    if (bEnd)
                        continue;

                    foreach (Match m in Regex.Matches(content, @"<[/]?[^<>]?>"))
                    {
                        string mInner = Regex.Replace(m.Value, "[/<>]", "");
                        if (!"sub:sup:image".Contains(mInner) && !Regex.IsMatch(mInner, "^[0-9]+$"))
                        {

                        }
                    }

                    content = content.Replace("&", "&amp;");
                    content = content.Replace("<", "&lt;");
                    content = content.Replace(">", "&gt;");

                    if (content.Contains("&lt;image&gt;"))
                    {
                        content = content.Replace("&lt;image&gt;", "<image>").Replace("&lt;/image&gt;", "</image>");
                        foreach (Match match in Regex.Matches(content, @"<image>[\s\S]+?</image>"))
                        {
                            content = content.Replace(match.Value, "<chemistry num=\"1\" id=\"chem001\"><img img-format=\"jpg\" file=\"chem001\" wi=\"\" he=\"\"></img></chemistry>");
                        }
                    }

                    content = content.Replace("&lt;sup&gt;", "<sup>").Replace("&lt;/sup&gt;", "</sup>").Replace("&lt;sub&gt;", "<sub>").Replace("&lt;/sub&gt;", "</sub>");

                    sw.WriteLine("<p num=\"" + iNum++ + "\">" + content + "</p>");
                }
                sw.WriteLine("\t\t</cn-holding>");
                sw.WriteLine("\t</cn-decision-detail>");
            }
            sw.WriteLine("</cn-appeal-decision>");
            sw.Flush();
            sw.Close();
            sw.Dispose();
        }

        public bool doCheckXml(string xmlPath)
        {
            bool bReturn = true;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.XmlResolver = null;
            try
            {
                xmlDoc.Load(xmlPath);

            }
            catch
            {
                bReturn = false;
            }
            return bReturn;
        }
    }
}
