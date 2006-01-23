/*  This is RTF to HTML converter, implemented as a text filter, generally.
    Copyright (C) 2003 Valentin Lavrinenko, vlavrinenko@users.sourceforge.net

    This library is free software; you can redistribute it and/or
    modify it under the terms of the GNU Lesser General Public
    License as published by the Free Software Foundation; either
    version 2.1 of the License, or (at your option) any later version.

    This library is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
    Lesser General Public License for more details.

    You should have received a copy of the GNU Lesser General Public
    License along with this library; if not, write to the Free Software
    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/


/*  ��� ��������� RTF � HTML, �������������, � ��������, ��� ��������� ������.
    Copyright (C) 2003 �������� ����������, vlavrinenko@users.sourceforge.net

    ������ ���������� �������� ��������� ����������� ������������. �� 
    ������ �������������� �� �/��� �������������� � ������������ � ��������� 
    ������ 2.1 ���� �� ������ ������ � ��������� ����� ������� ������ 
    ����������� ������������ �������� ������������� ���������� GNU, 
    �������������� Free Software Foundation.

    �� �������������� ��� ���������� � ������� �� ��, ��� ��� ����� ��� 
    ��������, ������ �� ������������� �� ��� ������� ��������, � ��� ����� 
    �������� ��������� ��������� ��� ������� � ����������� ��� ������������� 
    � ���������� �����. ��� ��������� ����� ��������� ���������� ������������ 
    �� ����������� ������������ ��������� ������������� ���������� GNU.

    ������ � ������ ����������� �� ������ ���� �������� ��������� ����������� 
    ������������ �������� ������������� ���������� GNU. ���� �� ��� �� 
    ��������, �������� �� ���� �� Free Software Foundation, Inc., 59 Temple 
    Place - Suite 330, Boston, MA 02111-1307, USA.
*/

#include "config.h"
#include "rtf_table.h"
#include "rtf_tools.h"
#include "rtf_keyword.h"
#include "fmt_opts.h"
#include <cctype>
#include <cstdlib>
#include <stdexcept>
#include <fstream>
#include <iostream>
#include <string>

//#include "dbg_iter.h"


int main(int argc, char **argv)
{
   try
   {
   if (argc==2)
   {
      // We do not want to use strcmp here because of compability issues
      std::string argv1(argv[1]);
      if (argv1=="-v" || argv1=="--version")
      {
         std::cout<<PACKAGE<<" version "<<PACKAGE_VERSION<<std::endl;
         return 0;
      }
      if (argv1=="-h" || argv1=="--help")
      {
         std::cout<<"Usage: "<<PACKAGE<<" [<rtf file> [<html file>]] "
                                      <<" - convert rtf to html"<<std::endl
                  <<"   or: "<<PACKAGE<<" -v|--version               "
                                      <<" - print current version number"<<std::endl
                  <<"   or: "<<PACKAGE<<" -h|--help                  "
                                      <<" - display this help message"<<std::endl;
         return 0;
      }
   }
   if (argc>3)
   {
      std::cerr<<"Too many arguments! See "<<PACKAGE<<" --help for options"<<std::endl;
      return 1;
   }
   std::istream *p_file_in=argc>1?new std::ifstream(argv[1]):&std::cin;
   std::ostream *p_file_out=argc>2?new std::ofstream(argv[2]):&std::cout;
   std::istream &file_in=*p_file_in;
   std::ostream &file_out=*p_file_out;
   file_in.exceptions(std::ios::failbit);
   file_out.exceptions(std::ios::failbit);
   std::istreambuf_iterator<char> f_buf_in(file_in.rdbuf());
   std::istreambuf_iterator<char> f_buf_in_end;
   std::string str_in;
   for (; f_buf_in != f_buf_in_end ; ++f_buf_in)
      str_in.push_back(*f_buf_in);
   std::string::iterator buf_in=str_in.begin(), buf_in_end=str_in.end();

   std::string html, par_html;
   bool bAsterisk=false;
   fo_stack foStack;
   formatting_options cur_options, new_options;

   /* CellDefs in rtf are really queer. We'll keep a list of them in main()
      and will give an iterator into this list to a row */
   table_cell_defs_list CellDefsList;
   table_cell_defs_list::iterator CurCellDefs;
   table_cell_def *tcdCurCellDef=new table_cell_def;
   table_cell *tcCurCell=new table_cell;
   table_row *trCurRow=new table_row;
   table *tblCurTable=new table;
   int iLastRowLeft=0, iLastRowHeight=0;

   bool bInTable=false;
   int iDocWidth=12240;
   int iMarginLeft=1800;
   while(buf_in!=buf_in_end)
   {
      switch (*buf_in)
      {
      case '\\':
      {
         rtf_keyword kw(++buf_in);
         if (kw.is_control_char())
            switch (kw.control_char())
            {
            case '\\': case '{': case '}':
               par_html+=kw.control_char();
               break;
            case '\'':
            {
               std::string stmp(1,*buf_in++);
               stmp+=*buf_in++;
               int code=std::strtol(stmp.c_str(), NULL, 16);
               switch (code)
               {
                  case 167:
                     par_html+="&bull;";
                     break;
                  case 188:
                     par_html+="&hellip;";
                     break;
                  default:
                     par_html+=(char)code;
               }
               break;
            }
            case '*':
               bAsterisk=true;
               break;
            case '~':
               par_html+="&nbsp;";
               break;
            }
         else //kw.is_control_char
            if (bAsterisk)
            {
               bAsterisk=false;
               skip_group(buf_in);
            }
            else
            {
               switch (kw.keyword())
               {
               case rtf_keyword::rkw_fonttbl: case rtf_keyword::rkw_filetbl: 
               case rtf_keyword::rkw_stylesheet: case rtf_keyword::rkw_info:
               case rtf_keyword::rkw_colortbl: case rtf_keyword::rkw_header: 
               case rtf_keyword::rkw_footer: case rtf_keyword::rkw_headerf: 
               case rtf_keyword::rkw_footerf: case rtf_keyword::rkw_pict:
               case rtf_keyword::rkw_object:
                  // we'll skip such groups
                  skip_group(buf_in);
                  break;
               // special characters
               case rtf_keyword::rkw_line: case rtf_keyword::rkw_softline:
                  par_html+="<br>";
                  break;
               case rtf_keyword::rkw_tab:
                  par_html+="&nbsp;&nbsp;";  // maybe, this can be done better
                  break;
               case rtf_keyword::rkw_enspace: case rtf_keyword::rkw_emspace:
                  par_html+="&nbsp;";
                  break;
               case rtf_keyword::rkw_qmspace:
                  par_html+="&thinsp;";
                  break;
               case rtf_keyword::rkw_endash:
                  par_html+="&ndash;";
                  break;
               case rtf_keyword::rkw_emdash:
                  par_html+="&mdash;";
                  break;
               case rtf_keyword::rkw_bullet:
                  par_html+="&bull;";
                  break;
               case rtf_keyword::rkw_lquote:
                  par_html+="&lsquo;";
                  break;
               case rtf_keyword::rkw_rquote:
                  par_html+="&rsquo;";
                  break;
               case rtf_keyword::rkw_ldblquote:
                  par_html+="&ldquo;";
                  break;
               case rtf_keyword::rkw_rdblquote:
                  par_html+="&rdquo;";
                  break;
               // paragraph formatting 
               case rtf_keyword::rkw_ql:
                  cur_options.papAlign=formatting_options::align_left;
                  break;
               case rtf_keyword::rkw_qr:
                  cur_options.papAlign=formatting_options::align_right;
                  break;
               case rtf_keyword::rkw_qc:
                  cur_options.papAlign=formatting_options::align_center;
                  break;
               case rtf_keyword::rkw_qj:
                  cur_options.papAlign=formatting_options::align_justify;
                  break;
               case rtf_keyword::rkw_fi:
                  cur_options.papFirst=(int)rint(kw.parameter()/20);
                  break;
               case rtf_keyword::rkw_li:
                  cur_options.papLeft=(int)rint(kw.parameter()/20);
                  break;
               case rtf_keyword::rkw_ri:
                  cur_options.papRight=(int)rint(kw.parameter()/20);
                  break;
               case rtf_keyword::rkw_sb:
                  cur_options.papBefore=(int)rint(kw.parameter()/20);
                  break;
               case rtf_keyword::rkw_sa:
                  cur_options.papAfter=(int)rint(kw.parameter()/20);
                  break;
               case rtf_keyword::rkw_pard:
                  cur_options.papBefore=cur_options.papAfter=0;
                  cur_options.papLeft=cur_options.papRight=0;
                  cur_options.papFirst=0;
                  cur_options.papAlign=formatting_options::align_left;
                  cur_options.papInTbl=false;
                  break;
               case rtf_keyword::rkw_par: case rtf_keyword::rkw_sect:
                  par_html=cur_options.get_par_str()+par_html;
                  if (cur_options.chpBold)
                     par_html+="</b>";
                  if (cur_options.chpItalic)
                     par_html+="</i>";
                  if (cur_options.chpUnderline)
                     par_html+="</u>";
                  if (cur_options.chpSup)
                     par_html+="</sup>";
                  if (cur_options.chpSub)
                     par_html+="</sub>";
                  par_html+="&nbsp;</p>";
                  if (!bInTable)
                  {
                     html+=par_html;
                  }
                  else 
                  {
                     if (cur_options.papInTbl)
                     {
                        tcCurCell->Text+=par_html;
                     }
                     else
                     {
                        html+=tblCurTable->make()+par_html;
                        bInTable=false;
                        delete tblCurTable;
                        tblCurTable=new table;
                     }
                  }
                  par_html="";
                  if (cur_options.chpBold)
                     par_html+="<b>";
                  if (cur_options.chpItalic)
                     par_html+="<i>";
                  if (cur_options.chpUnderline)
                     par_html+="<u>";
                  if (cur_options.chpSup)
                     par_html+="<sup>";
                  if (cur_options.chpSub)
                     par_html+="<sub>";
               break;
               // character formatting
               case rtf_keyword::rkw_super:
                  cur_options.chpSup=!(kw.parameter()==0);
                  insert_char_option(par_html, "sup", cur_options.chpSup);
                  break;
               case rtf_keyword::rkw_sub:
                  cur_options.chpSub=!(kw.parameter()==0);
                  insert_char_option(par_html, "sub", cur_options.chpSub);
                  break;
               case rtf_keyword::rkw_b:
                  cur_options.chpBold=!(kw.parameter()==0);
                  insert_char_option(par_html, "b", cur_options.chpBold);
                  break;
               case rtf_keyword::rkw_i:
                  cur_options.chpItalic=!(kw.parameter()==0);
                  insert_char_option(par_html, "i", cur_options.chpItalic);
                  break;
               case rtf_keyword::rkw_ul:
                  cur_options.chpUnderline=!(kw.parameter()==0);
                  insert_char_option(par_html, "u", cur_options.chpUnderline);
                  break;
               case rtf_keyword::rkw_ulnone:
                  cur_options.chpUnderline=false;
                  insert_char_option(par_html, "u", cur_options.chpUnderline);
                  break;
               case rtf_keyword::rkw_plain:
                  if (cur_options.chpBold)
                  {
                     cur_options.chpBold=false;
                     insert_char_option(par_html, "b", false);
                  }
                  if (cur_options.chpItalic)
                  {
                     cur_options.chpItalic=false;
                     insert_char_option(par_html, "i", false);
                  }
                  if (cur_options.chpUnderline)
                  {
                     cur_options.chpUnderline=false;
                     insert_char_option(par_html, "u", false);
                  }
                  if (cur_options.chpSup)
                  {
                     cur_options.chpSup=false;
                     insert_char_option(par_html, "sup", false);
                  }
                  if (cur_options.chpSub)
                  {
                     cur_options.chpSub=false;
                     insert_char_option(par_html, "sub", false);
                  }
                  break;
               // table formatting
               case rtf_keyword::rkw_intbl:
                  cur_options.papInTbl=true;
                  break;
               case rtf_keyword::rkw_trowd: 
                  CurCellDefs=CellDefsList.insert(CellDefsList.end(), 
                                                  table_cell_defs());
               case rtf_keyword::rkw_row:
                  if (!trCurRow->Cells.empty())
                  {
                     trCurRow->CellDefs=CurCellDefs;
                     if (trCurRow->Left==-1000)
                        trCurRow->Left=iLastRowLeft;
                     if (trCurRow->Height==-1000)
                        trCurRow->Height=iLastRowHeight;
                     tblCurTable->push_back(trCurRow);
                     trCurRow=new table_row;
                  }
                  bInTable=true;
                  break;
               case rtf_keyword::rkw_cell:
                  par_html=cur_options.get_par_str()+par_html;
                  if (cur_options.chpBold)
                     par_html+="</b>";
                  if (cur_options.chpItalic)
                     par_html+="</i>";
                  if (cur_options.chpUnderline)
                     par_html+="</u>";
                  if (cur_options.chpSup)
                     par_html+="</sup>";
                  if (cur_options.chpSub)
                     par_html+="</sub>";
                  par_html+=+"</p>";
                  tcCurCell->Text+=par_html;
                  par_html="";
                  if (cur_options.chpBold)
                     par_html+="<b>";
                  if (cur_options.chpItalic)
                     par_html+="<i>";
                  if (cur_options.chpUnderline)
                     par_html+="<u>";
                  if (cur_options.chpSup)
                     par_html+="<sup>";
                  if (cur_options.chpSub)
                     par_html+="<sub>";
                  trCurRow->Cells.push_back(tcCurCell);
                  tcCurCell=new table_cell;
                  break;
               case rtf_keyword::rkw_cellx:
                  tcdCurCellDef->Right=kw.parameter();
                  CurCellDefs->push_back(tcdCurCellDef);
                  tcdCurCellDef=new table_cell_def;
                  break;
               case rtf_keyword::rkw_trleft:
                  trCurRow->Left=kw.parameter();
                  iLastRowLeft=kw.parameter();
                  break;
               case rtf_keyword::rkw_trrh:
                  trCurRow->Height=kw.parameter();
                  iLastRowHeight=kw.parameter();
                  break;
               case rtf_keyword::rkw_clvmgf:
                  tcdCurCellDef->FirstMerged=true;
                  break;
               case rtf_keyword::rkw_clvmrg:
                  tcdCurCellDef->Merged=true;
                  break;
               case rtf_keyword::rkw_clbrdrb:
                  tcdCurCellDef->BorderBottom=true;
                  tcdCurCellDef->ActiveBorder=&(tcdCurCellDef->BorderBottom);
                  break;
               case rtf_keyword::rkw_clbrdrt:
                  tcdCurCellDef->BorderTop=true;
                  tcdCurCellDef->ActiveBorder=&(tcdCurCellDef->BorderTop);
                  break;
               case rtf_keyword::rkw_clbrdrl:
                  tcdCurCellDef->BorderLeft=true;
                  tcdCurCellDef->ActiveBorder=&(tcdCurCellDef->BorderLeft);
                  break;
               case rtf_keyword::rkw_clbrdrr:
                  tcdCurCellDef->BorderRight=true;
                  tcdCurCellDef->ActiveBorder=&(tcdCurCellDef->BorderRight);
                  break;
               case rtf_keyword::rkw_brdrnone:
                  if (tcdCurCellDef->ActiveBorder!=NULL)
                  {
                     *(tcdCurCellDef->ActiveBorder)=false;
                  }
                  break;
               case rtf_keyword::rkw_clvertalt:
                  tcdCurCellDef->VAlign=table_cell_def::valign_top;
                  break;
               case rtf_keyword::rkw_clvertalc:
                  tcdCurCellDef->VAlign=table_cell_def::valign_center;
                  break;
               case rtf_keyword::rkw_clvertalb:
                  tcdCurCellDef->VAlign=table_cell_def::valign_bottom;
                  break;
               // page formatting
               case rtf_keyword::rkw_paperw:
                  iDocWidth=kw.parameter();
                  break;
               case rtf_keyword::rkw_margl:
                  iMarginLeft=kw.parameter();
                  break;
               }
            }
         break;
      }
      case '{':
         // perform group opening actions here
         foStack.push(cur_options);
         ++buf_in;
         break;
      case '}':
         // perform group closing actions here
         new_options=foStack.top();
         foStack.pop();
         if (cur_options.chpBold!=new_options.chpBold)
         {
            insert_char_option(par_html, "b", new_options.chpBold);
         }
         if (cur_options.chpItalic!=new_options.chpItalic)
         {
            insert_char_option(par_html, "i", new_options.chpItalic);
         }
         if (cur_options.chpUnderline!=new_options.chpUnderline)
         {
            insert_char_option(par_html, "u", new_options.chpUnderline);
         }
         if (cur_options.chpSup!=new_options.chpSup)
         {
            insert_char_option(par_html, "sup", new_options.chpSup);
         }
         if (cur_options.chpSub!=new_options.chpSub)
         {
            insert_char_option(par_html, "sub", new_options.chpSub);
         }

         memcpy(&cur_options, &new_options, sizeof(cur_options));
         ++buf_in;
         break;
      case 13: case 10:
         ++buf_in;
         break;
      case '<':
         par_html+="&lt;";
         ++buf_in;
         break;
      case '>':
         par_html+="&gt;";
         ++buf_in;
         break;
      default:
         par_html+=*buf_in++;
      }
   }
   file_out<<"<html><head></head><body><div style=\"width:";
   file_out<<rint((iDocWidth/20));
   file_out<<"pt;";
   file_out<<"padding-left=";
   file_out<<rint(iMarginLeft/20);
   file_out<<"pt\">";
   file_out<<html;
   file_out<<"</div></body></html>";
   if (argc>1)
      delete p_file_in;
   if (argc>2)
      delete p_file_out;
   delete tcCurCell;
   delete trCurRow;
   delete tblCurTable;
   delete tcdCurCellDef;
   return 0;
   }
   catch (std::exception &e)
   {
      std::cerr<<"Error: "<<e.what()<<std::endl;
   }
   catch (...)
   {
      std::cerr<<"Something really bad happened!";
   }
}
