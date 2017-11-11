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


/*  Это конвертер RTF в HTML, реализованный, в принципе, как текстовый фильтр.
    Copyright (C) 2003 Валентин Лавриненко, vlavrinenko@users.sourceforge.net

    Данная библиотека является свободным программным обеспечением. Вы 
    вправе распространять ее и/или модифицировать в соответствии с условиями 
    версии 2.1 либо по вашему выбору с условиями более поздней версии 
    Стандартной Общественной Лицензии Ограниченного Применения GNU, 
    опубликованной Free Software Foundation.

    Мы распространяем эту библиотеку в надежде на то, что она будет вам 
    полезной, однако НЕ ПРЕДОСТАВЛЯЕМ НА НЕЕ НИКАКИХ ГАРАНТИЙ, в том числе 
    ГАРАНТИИ ТОВАРНОГО СОСТОЯНИЯ ПРИ ПРОДАЖЕ и ПРИГОДНОСТИ ДЛЯ ИСПОЛЬЗОВАНИЯ 
    В КОНКРЕТНЫХ ЦЕЛЯХ. Для получения более подробной информации ознакомьтесь 
    со Стандартной Общественной Лицензией Ограниченного Применений GNU.

    Вместе с данной библиотекой вы должны были получить экземпляр Стандартной 
    Общественной Лицензии Ограниченного Применения GNU. Если вы его не 
    получили, сообщите об этом во Free Software Foundation, Inc., 59 Temple 
    Place - Suite 330, Boston, MA 02111-1307, USA.
*/

#include "config.h"
#include "rtf_table.h"
#include "rtf_tools.h"
#include "rtf_keyword.h"
#include "fmt_opts.h"
#include <cstdlib>
#include <stdexcept>
#include <string.h>
#include <stdint.h>
#include <fstream>
#include <iostream>
#include <string>

//#include "dbg_iter.h"

enum {
	UTF8ERR_INVALID=-1,
	UTF8ERR_NO_ROOM=-2
};

#ifndef UNICODE_BOM
#define UNICODE_BOM 0xFEFF
#endif

typedef char utf8_t;
typedef uint16_t utf16_t;

/* [doc] utf8_encode
 *
 * Encode unicode character 'code' as UTF-8 ASCII string
 *
 * Parameters:
 *
 *    p = pointer to a char* pointer where the output will be written.
 *        on return the pointer will have been updated.
 *
 *    fence = first byte past the end of the buffer. the function will not
 *            write the output and update the pointer if doing so would bring
 *            it past this point, in order to prevent buffer overrun and
 *            possible memory corruption issues
 *
 *    c = unicode character to encode
 *
 * Warning:
 *
 *    Remember that one encoded UTF-8 character can occupy anywhere between
 *    1 to 4 bytes. Do not assume one byte = one char. Use the fence pointer
 *    to prevent buffer overruns and to know when the buffer should be
 *    emptied and refilled.
 * 
 */
int utf8_encode(char **ptr,char *fence,uint32_t code) {
	int uchar_size=1;
	char *p = *ptr;

	if (!p) return UTF8ERR_NO_ROOM;
	if (code >= (uint32_t)0x80000000UL) return UTF8ERR_INVALID;
	if (p >= fence) return UTF8ERR_NO_ROOM;

	if (code >= 0x4000000) uchar_size = 6;
	else if (code >= 0x200000) uchar_size = 5;
	else if (code >= 0x10000) uchar_size = 4;
	else if (code >= 0x800) uchar_size = 3;
	else if (code >= 0x80) uchar_size = 2;

	if ((p+uchar_size) > fence) return UTF8ERR_NO_ROOM;

	switch (uchar_size) {
		case 1:	*p++ = (char)code;
			break;
		case 2:	*p++ = (char)(0xC0 | (code >> 6));
			*p++ = (char)(0x80 | (code & 0x3F));
			break;
		case 3:	*p++ = (char)(0xE0 | (code >> 12));
			*p++ = (char)(0x80 | ((code >> 6) & 0x3F));
			*p++ = (char)(0x80 | (code & 0x3F));
			break;
		case 4:	*p++ = (char)(0xF0 | (code >> 18));
			*p++ = (char)(0x80 | ((code >> 12) & 0x3F));
			*p++ = (char)(0x80 | ((code >> 6) & 0x3F));
			*p++ = (char)(0x80 | (code & 0x3F));
			break;
		case 5:	*p++ = (char)(0xF8 | (code >> 24));
			*p++ = (char)(0x80 | ((code >> 18) & 0x3F));
			*p++ = (char)(0x80 | ((code >> 12) & 0x3F));
			*p++ = (char)(0x80 | ((code >> 6) & 0x3F));
			*p++ = (char)(0x80 | (code & 0x3F));
			break;
		case 6:	*p++ = (char)(0xFC | (code >> 30));
			*p++ = (char)(0x80 | ((code >> 24) & 0x3F));
			*p++ = (char)(0x80 | ((code >> 18) & 0x3F));
			*p++ = (char)(0x80 | ((code >> 12) & 0x3F));
			*p++ = (char)(0x80 | ((code >> 6) & 0x3F));
			*p++ = (char)(0x80 | (code & 0x3F));
			break;
	};

	*ptr = p;
	return 0;
}

static std::string ansi2utf(char c,std::string &font) {
    // hack: some WinHelp documents have sample code with spaces to indent them.
    // if the font is Courier New, convert spaces to &nbsp; to prevent the browser from collapsing them.
    if (c == ' ') {
        if (!strcasecmp(font.c_str(),"courier") || !strcasecmp(font.c_str(),"courier new"))
            return "&nbsp;";
    }

    // the reason this is a function is so that later revisions of the code
    // can support other legacy non-unicode encodings should the user run
    // this program against non-latin-1 encodings.
    //
    // also, so that we can add command line switches later for purists who
    // do NOT want us converting charsets.
    if ((unsigned char)c < (unsigned char)0x80) {
        std::string r;
        r = c; // FIXME: there's no constructor for char?
        return r;
    }

    {
        char tmp[32]={0},*w=tmp,*f=tmp+sizeof(tmp)-1;

        if (!strcasecmp(font.c_str(),"SYMBOL")) {
            // Windows Symbol font to UTF-8 conversion.
            // This isn't complete by far, but it's enough to convert the most commonly used ones in WinHelp files.
            switch ((unsigned char)c) {
                case 180: // Multiplication symbol
                    utf8_encode(&w,f,0x00D7);
                    break;
                case 210: // Registered symbol
                    utf8_encode(&w,f,0x00AE);
                    break;
                case 211: // Copyright symbol
                    utf8_encode(&w,f,0x00A9);
                    break;
                case 212: // Trademark symbol
                    utf8_encode(&w,f,0x2122);
                    break;
                default:
                    fprintf(stderr,"FIXME: Windows Symbol font code 0x%02x not implemented\n",(unsigned char)c);
                    break;
            };
        }
        else {
            // ANSI (Latin-1)
            switch ((unsigned char)c) {
                case 0x86: // dagger
                    utf8_encode(&w,f,0x2020);
                    break;
                case 0x87: // double dagger
                    utf8_encode(&w,f,0x2021);
                    break;
                case 0x93: // left double quotation mark
                    utf8_encode(&w,f,0x201C);
                    break;
                case 0x94: // right double quotation mark
                    utf8_encode(&w,f,0x201D);
                    break;
                case 0x96: // en dash
                    utf8_encode(&w,f,0x2013);
                    break;
                case 0x97: // em dash
                    utf8_encode(&w,f,0x2014);
                    break;
                case 0xA0: // NBSP
                    utf8_encode(&w,f,0x00A0);
                    break;
                case 0xA9: // copyright
                    utf8_encode(&w,f,0x00A9);
                    break;
                case 0xB7: // dot
                    utf8_encode(&w,f,0x00B7);
                    break;
                default:
                    fprintf(stderr,"FIXME: ANSI to UTF-8 for code 0x%02x not implemented\n",(unsigned char)c);
                    break;
            }
        }
    
        return (const char*)tmp;
    }
}

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
   colorvect colortbl;
   fontmap fonttbl;
   std::string title;

   bool bAsterisk=false;
   fo_stack foStack;
   formatting_options cur_options;
   std::string html;
   html_text par_html(cur_options);

   /* CellDefs in rtf are really queer. We'll keep a list of them in main()
      and will give an iterator into this list to a row */
   table_cell_defs_list CellDefsList;
   table_cell_defs_list::iterator CurCellDefs;
   table_cell_def *tcdCurCellDef=new table_cell_def;
   table_cell *tcCurCell=new table_cell;
   table_row *trCurRow=new table_row;
   table *tblCurTable=new table;
   int iLastRowLeft=0, iLastRowHeight=0;
   std::string t_str;

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
         {
            switch (kw.control_char())
            {
            case '\\': case '{': case '}':
               par_html.write(kw.control_char());
               break;
            case '\'':
               par_html.write(ansi2utf(char_by_code_noesc(buf_in),cur_options.chpFont.name));
               break;
            case '*':
               bAsterisk=true;
               break;
            case '~':
               par_html.write("&nbsp;");
               break;
            }
         }
         else //kw.is_control_char
         {
            if (bAsterisk)
            {
               bAsterisk=false;
               skip_group(buf_in);
               cur_options=foStack.top();
               foStack.pop();
            }
            else
            {
               switch (kw.keyword())
               {
               case rtf_keyword::rkw_filetbl: 
               case rtf_keyword::rkw_stylesheet:
               case rtf_keyword::rkw_header: 
               case rtf_keyword::rkw_footer: case rtf_keyword::rkw_headerf: 
               case rtf_keyword::rkw_footerf: case rtf_keyword::rkw_pict:
               case rtf_keyword::rkw_object:
                  // we'll skip such groups
                  skip_group(buf_in);
                  break;
               // document title
               case rtf_keyword::rkw_info: 
               {
                  int depth=1;
                  bool in_title=false;
                  while (depth>0)
                  {
//                     std::cout<<std::string(buf_in).substr(0,20)<<"\t"<<depth<<std::endl;
                     switch (*buf_in)
                     {
                     case '\\':
                     {
                        rtf_keyword kw(++buf_in);
                        if (in_title && kw.is_control_char() && kw.control_char() == '\'')
                           title += ansi2utf(char_by_code_noesc(buf_in),cur_options.chpFont.name);
                        else if (kw.keyword()==rtf_keyword::rkw_title)
                           in_title=true;
                        break;
                     }
                     case '{': 
                        ++depth; ++buf_in;
                        break;
                     case '}':
                        --depth; ++buf_in; 
                        in_title=false;
                        break;
                     default:
                        if (in_title)
                            title += *buf_in;
                        ++buf_in;
                        break;
                     }
                  }
                  break;
               }
               // color table
               case rtf_keyword::rkw_colortbl:
               {
                  color clr;
                  while (*buf_in!='}')
                  {
                     switch (*buf_in)
                     {
                     case '\\':
                     {
                        rtf_keyword kw(++buf_in);
                        switch (kw.keyword())
                        {
                        case rtf_keyword::rkw_red:
                           clr.r=kw.parameter();
                           break;
                        case rtf_keyword::rkw_green:
                           clr.g=kw.parameter();
                           break;
                        case rtf_keyword::rkw_blue:
                           clr.b=kw.parameter();
                           break;
                        }
                        break;
                     }
                     case ';':
                        colortbl.push_back(clr);
                        ++buf_in;
                        break;
                     default:
                        ++buf_in;
                        break;
                     }
                  }
                  ++buf_in;
                  break;
               }
               // font table
               case rtf_keyword::rkw_fonttbl: 
               {
                  font fnt;
                  int font_num;
                  bool full_name=false;
                  bool in_font=false;
                  while (! (*buf_in=='}' && !in_font))
                  {
                     switch (*buf_in)
                     {
                     case '\\':
                     {
                        rtf_keyword kw(++buf_in);
                        if (kw.is_control_char() && kw.control_char()=='*')
                           skip_group(buf_in);
                        else
                           switch (kw.keyword())
                           {
                           case rtf_keyword::rkw_f:
                              font_num=kw.parameter();
                              break;
                           case rtf_keyword::rkw_fprq:
                              fnt.pitch=kw.parameter();
                              break;
                           case rtf_keyword::rkw_fcharset:
                              fnt.charset=kw.parameter();
                              break;
                           case rtf_keyword::rkw_fnil:
                              fnt.family=font::ff_none;
                              break;
                           case rtf_keyword::rkw_froman:
                              fnt.family=font::ff_serif;
                              break;
                           case rtf_keyword::rkw_fswiss:
                              fnt.family=font::ff_sans_serif;
                              break;
                           case rtf_keyword::rkw_fmodern:
                              fnt.family=font::ff_monospace;
                              break;
                           case rtf_keyword::rkw_fscript:
                              fnt.family=font::ff_cursive;
                              break;
                           case rtf_keyword::rkw_fdecor:
                              fnt.family=font::ff_fantasy;
                              break;
                           }
                        break;
                     }
                     case '{':
                        in_font=true;
                        ++buf_in;
                        break;
                     case '}':
                        in_font=false;
                        fonttbl.insert(std::make_pair(font_num, fnt));
                        fnt=font();
                        full_name=false;
                        ++buf_in;
                        break;
                     case ';':
                        full_name=true;
                        ++buf_in;
                        break;
                     default:
                        if (!full_name && in_font)
                           fnt.name+=*buf_in;
                        ++buf_in;
                        break;
                     }
                  }
                  ++buf_in;
                  break;
               }
               // special characters
               case rtf_keyword::rkw_line: case rtf_keyword::rkw_softline:
                  par_html.write("<br>");
                  break;
               case rtf_keyword::rkw_tab:
                  par_html.write("&nbsp;&nbsp;");  // maybe, this can be done better
                  break;
               case rtf_keyword::rkw_enspace: case rtf_keyword::rkw_emspace:
                  par_html.write("&nbsp;");
                  break;
               case rtf_keyword::rkw_qmspace:
                  par_html.write("&thinsp;");
                  break;
               case rtf_keyword::rkw_endash:
                  par_html.write("&ndash;");
                  break;
               case rtf_keyword::rkw_emdash:
                  par_html.write("&mdash;");
                  break;
               case rtf_keyword::rkw_bullet:
                  par_html.write("&bull;");
                  break;
               case rtf_keyword::rkw_lquote:
                  par_html.write("&lsquo;");
                  break;
               case rtf_keyword::rkw_rquote:
                  par_html.write("&rsquo;");
                  break;
               case rtf_keyword::rkw_ldblquote:
                  par_html.write("&ldquo;");
                  break;
               case rtf_keyword::rkw_rdblquote:
                  par_html.write("&rdquo;");
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
                  t_str=cur_options.get_par_str()+par_html.str()
                        +"&nbsp;"+par_html.close()+"</p>\n";
                  if (!bInTable)
                  {
                     html+=t_str;
                  }
                  else 
                  {
                     if (cur_options.papInTbl)
                     {
                        tcCurCell->Text+=t_str;
                     }
                     else
                     {
                        html+=tblCurTable->make()+t_str;
                        bInTable=false;
                        delete tblCurTable;
                        tblCurTable=new table;
                     }
                  }
                  par_html.clear();
                  break;
               // character formatting
               case rtf_keyword::rkw_super:
                  cur_options.chpVAlign=
                     kw.parameter()==0?formatting_options::va_normal
                                      :formatting_options::va_sup;
                  break;
               case rtf_keyword::rkw_sub:
                  cur_options.chpVAlign=
                     kw.parameter()==0?formatting_options::va_normal
                                      :formatting_options::va_sub;
                  break;
               case rtf_keyword::rkw_v:
                  cur_options.chpHidden=true;
                  break;
               case rtf_keyword::rkw_b:
                  cur_options.chpBold=!(kw.parameter()==0);
                  break;
               case rtf_keyword::rkw_caps:
                  cur_options.chpAllCaps=!(kw.parameter()==0);
                  break;
               case rtf_keyword::rkw_i:
                  cur_options.chpItalic=!(kw.parameter()==0);
                  break;
               case rtf_keyword::rkw_ul:
                  cur_options.chpUnderline=!(kw.parameter()==0);
                  break;
               case rtf_keyword::rkw_ulnone:
                  cur_options.chpUnderline=false;
                  break;
               case rtf_keyword::rkw_fs:
                  cur_options.chpFontSize=kw.parameter();
                  break;
               case rtf_keyword::rkw_dn:
                  cur_options.chpVShift = kw.parameter() == -1 ? 6 : kw.parameter();
                  break;
               case rtf_keyword::rkw_up:
                  cur_options.chpVShift = kw.parameter() == -1 ? -6 : -kw.parameter();
                  break;
               case rtf_keyword::rkw_cf:
                  cur_options.chpFColor=colortbl[kw.parameter()];
                  break;
               case rtf_keyword::rkw_cb:
                  cur_options.chpBColor=colortbl[kw.parameter()];
                  break;
               case rtf_keyword::rkw_highlight:
                  cur_options.chpHighlight=kw.parameter();
                  break;
               case rtf_keyword::rkw_f:
                  cur_options.chpFont=fonttbl[kw.parameter()];
                  break;
               case rtf_keyword::rkw_plain:
                  cur_options.chpBold=cur_options.chpAllCaps
                    =cur_options.chpItalic=cur_options.chpUnderline=false;
                  cur_options.chpVAlign=formatting_options::va_normal;
                  cur_options.chpFontSize=cur_options.chpHighlight=0;
                  cur_options.chpFColor=cur_options.chpBColor=color();
                  cur_options.chpFont=font();
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
                  t_str=cur_options.get_par_str()+par_html.str()
                        +"&nbsp;"+par_html.close()+"</p>\n";
                  tcCurCell->Text+=t_str;
                  par_html.clear();
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
         cur_options=foStack.top();
         foStack.pop();
         ++buf_in;
         break;
      case 13: case 10:
         ++buf_in;
         break;
      case '<':
         par_html.write("&lt;");
         ++buf_in;
         break;
      case '>':
         par_html.write("&gt;");
         ++buf_in;
         break;
/*      case ' ':
         par_html.write("&ensp;");
         ++buf_in;
         break;*/
      default:
         par_html.write(ansi2utf(*buf_in++,cur_options.chpFont.name));
      }
   }
   file_out<<"<html>\n<head>\n"
           <<"<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />\n"
           <<"<STYLE type=\"text/css\">\nbody {padding-left:"
           <<rint(iMarginLeft/20)<<"pt;width:"<<rint((iDocWidth/20))<<"pt;}\n"
           <<"p {margin-top:0pt;margin-bottom:0pt}\n"<<formatting_options::get_styles()<<"</STYLE>\n"
           <<"<title>"<<title<<"</title>\n</head>\n"
           <<"<body>\n"<<html<<"</body>\n</html>";
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

