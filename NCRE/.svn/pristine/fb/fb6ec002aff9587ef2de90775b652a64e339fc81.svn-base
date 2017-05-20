using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLL;
using Model;
using ppt = Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Microsoft.Win32;
using SHDocVw;
using pp1 = Microsoft.Office.Interop.PowerPoint.Presentation;

namespace NCRE学生考试端V1._0.PPT操作题类
{
    public class PptQuestionFlag
    {
         private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();
        #region 试卷A2015年12月11日17:34:16
        public void PptSwitchQuestionFlagA(PptQuestionEntity pptquestion)
        {
            //字体颜色
            PptFontColor pptfontcolor = new PptFontColor();
            //占位符
            PptPlaceholder pptplaceholder = new PptPlaceholder();
            //移动幻灯片
            PptMoveA pptmovea = new PptMoveA();
            //超链接
            PptHyperlinkA ppthyperlinka = new PptHyperlinkA();
            //添加幻灯片
            PptNewA pptnewa = new PptNewA();
            //幻灯片版式
            PptFormatA pptformata = new PptFormatA();
            //艺术字样式
            PptWordartA pptwordarta = new PptWordartA();
            //插入字
            PptArtWordTextA pptwordarttexta = new PptArtWordTextA();
            //字体
            PptFontNameA pptfontnamea = new PptFontNameA();
            //字号
            PptSizeA pptsizea = new PptSizeA();
            //加粗
            PptBoldA pptbolda = new PptBoldA();
            //艺术字形状
            PptWordartTypeA pptwordarttypea = new PptWordartTypeA();
            //切换效果
            PptActionTypeA pptactiontypea = new PptActionTypeA();
            //切换时间
            PptActionTypeTimeA pptactiontypetimea = new PptActionTypeTimeA();
            //主题
            PptThemeA pptthemea = new PptThemeA();
            pptquestion.PaperType = "A";
            DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                switch (questionflag)
                {
                   
                    case "字体颜色":
                        pptfontcolor.FontColor(pptquestion);
                        break;
                    case "删除占位符":
                        pptplaceholder.DelPlaceholder(pptquestion);
                        break;
                    case "移动幻灯片":
                        pptmovea.Move(pptquestion);
                        break;
                    case "超链接":
                        ppthyperlinka.ppthyperlink(pptquestion);
                        break;
                    case "添加幻灯片":
                        pptnewa.New (pptquestion);
                        break;
                    case "幻灯片版式":
                        pptformata.actionType(pptquestion);
                        break;
                    case "艺术字样式":
                        pptwordarta.actionType(pptquestion);
                        break;
                    case "插入字":
                        pptwordarttexta.WordtextA(pptquestion);
                        break;
                    case "字体":
                        pptfontnamea.FontName(pptquestion);
                        break;
                    case "字号":
                        pptsizea.size(pptquestion);
                        break;
                    case "加粗":
                        pptbolda.Bold(pptquestion);
                        break;
                    case "艺术字形状":
                        pptwordarttypea.wordType(pptquestion);
                        break;
                    case "切换效果":
                        pptactiontypea.actionType(pptquestion);
                        break;
                    case "切换时间":
                        pptactiontypetimea.pptactiontypetime(pptquestion);
                        break;
                    case "主题":
                        pptthemea.Theme(pptquestion);
                        break;

                }
            }

        }
        #endregion

        #region 试卷B2015年12月11日17:51:01
        public void PptSwitchQuestionFlagB(PptQuestionEntity pptquestion)
        {
            //标题字体
            PptFontNameTitleB pptfontnameTitleb = new PptFontNameTitleB();
            //标题字号
            PptSizeTitleB pptsizetitleb = new PptSizeTitleB();
            //字体颜色
            PptFontColorB pptfontcolorb = new PptFontColorB();
            
            //占位符
            PptPlaceholderB pptplaceholderb = new PptPlaceholderB();
            //移动幻灯片
            PptMoveB pptmoveb = new PptMoveB();
            //动画效果
            PptAnimationEffectB pptanimationeffectb = new PptAnimationEffectB();
            //添加幻灯片
            PptNewB pptnewb = new PptNewB();
            //幻灯片版式
            PptFormatB pptformatb = new PptFormatB();
            //艺术字样式
            PptWordartB pptwordartb = new PptWordartB();
            //插入字
            PptArtWordTextB pptwordarttextb = new PptArtWordTextB();
            //艺术字字体
            PptArtFontNameB pptartfontnameb = new PptArtFontNameB();
            //字号
            PptSizeB pptsizeb = new PptSizeB();
            //加粗
            PptBoldB pptboldb = new PptBoldB();
            //艺术字形状
            PptWordartTypeB pptwordarttypeb = new PptWordartTypeB();
            //切换效果
            PptActionTypeB pptactiontypeb = new PptActionTypeB();
            //切换时间
            PptActionTypeTimeB pptactiontypetimeb = new PptActionTypeTimeB();
            //主题
            PptThemeB pptthemea = new PptThemeB();
           
          DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);
          pptquestion.PaperType = "B";
          for (int i = 0; i < dt.Rows.Count; i++)
          {
              string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
              switch (questionflag)
              {
                  case "标题字体":

                      pptfontnameTitleb.FontName(pptquestion);
                      break;
                  case "标题字号":
                      pptsizetitleb.size(pptquestion);
                      break;
                  case "字体颜色":
                      pptfontcolorb.FontColor(pptquestion);
                      break;
                  case "删除占位符":
                      pptplaceholderb.DelPlaceholder(pptquestion);
                      break;
                  case "移动幻灯片":
                      pptmoveb.Move(pptquestion);
                      break;
                  case "动画效果":
                      pptanimationeffectb.PptAnimationEffect(pptquestion);
                      break;
                  case "添加幻灯片":
                      pptnewb.Add(pptquestion);
                      break;
                  case "幻灯片版式":
                      pptformatb.format(pptquestion);
                      break;
                  case "艺术字样式":
                      pptwordartb.WordartType(pptquestion);
                      break;
                  case "插入字":
                      pptwordarttextb.Artword(pptquestion);
                      break;
                  case "艺术字字体":
                      pptartfontnameb.FontName(pptquestion);
                      break;
                  case "字号":
                      pptsizeb.size(pptquestion);
                      break;
                  case "加粗":
                      pptboldb.bold(pptquestion);
                      break;
                  case "艺术字形状":
                      pptwordarttypeb.shape(pptquestion);
                      break;
                  case "切换效果":
                      pptactiontypeb.actionType(pptquestion);
                      break;
                  case "切换时间":
                      pptactiontypetimeb.actionTypeTime(pptquestion);
                      break;
                  case "主题":
                      pptthemea.themeB(pptquestion);
                      break;
              }
          }
        }
        #endregion

        #region 试卷C2015年12月11日16:03:57没有问题
        public void PptSwitchQuestionFlagC(PptQuestionEntity pptquestion)
        {
            //主题
            PptThemeC pptthemec = new PptThemeC();
            //移动幻灯片
            PptMoveC pptmovec = new PptMoveC();
            //删除幻灯片
            PptDelC pptdelc = new PptDelC();


            //主标题字体
            PptFontNameC pptfontnamec = new PptFontNameC();
            //主标题字号
            PptTitleSizeC pptsizec = new PptTitleSizeC(); 
            //主标题加粗
            PptBoldC pptboldc = new PptBoldC();
            //副标题字体
            PptFontSubTitleC pptfontsubtitlec=new PptFontSubTitleC ();
            //副标题字号
            PptSizeSubTitleSizeC pptsubtitlesizec = new PptSizeSubTitleSizeC();


           //切换方式
            PptActionTypeC pptactiontypec = new PptActionTypeC();
            //切换时间
            PptActionTypeTimeC pptactiontypetimec = new PptActionTypeTimeC();
            //图片大小
            PptPictureSizeC pptpicturesizec = new PptPictureSizeC();
            //动画效果PptAdvanceTimeC
            PptAnimationEffectC pptanimationeffectc = new PptAnimationEffectC();
            //动画时间
            PptAnimationEffectTimeC pptanimationeffecttimec = new PptAnimationEffectTimeC();
            pptquestion.PaperType = "C";
            DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                switch (questionflag)
                {
                    case "主题":
                        pptthemec.actionType(pptquestion);
                        break;
                    case "移动幻灯片":
                        pptmovec.Move(pptquestion);
                        break;
                    case "删除幻灯片":
                        pptdelc.Move(pptquestion);
                        break;
                    case "标题字体":
                        pptfontnamec.FontName(pptquestion);
                        break;
                    case "标题字号":
                        pptsizec.size(pptquestion);
                        break;
                    case "加粗":
                        pptboldc.BoldC(pptquestion);
                        break;
                    case "字体":
                        pptfontsubtitlec.FontName(pptquestion);
                        break;
                    case "字号":
                        pptsubtitlesizec.size(pptquestion);
                        break;

                    case "切换效果":
                        pptactiontypec.actionType(pptquestion);
                        break;
                    case "切换时间":
                        pptactiontypetimec.actionType(pptquestion);
                        break;
                    case "图片大小":
                        pptpicturesizec.pptpicturesize(pptquestion);
                        break;
                    case "动画效果":
                        pptanimationeffectc.PptAnimationEffect(pptquestion);
                        break;
                    case "动画时间":
                        pptanimationeffecttimec.PptAnimationEffect(pptquestion);
                        break;
                   
                }
            }
        }
        #endregion

        #region  试卷D2015年12月11日16:26:41
        public void SwitchQuestionFlagD(PptQuestionEntity pptquestion)
        {

            PptNewD pptnewd = new PptNewD();
            PptPictureSizeD pptpicturesized = new PptPictureSizeD();
            PptHyperlinkD ppthyperlinkd = new PptHyperlinkD();
            PptAnimationEffectD pptanimationeffectd = new PptAnimationEffectD();
            PptAnimationEffectTimeD pptanimationeffecttimed = new PptAnimationEffectTimeD();
            PptActionTypeD pptactiontyped = new PptActionTypeD();
            PptSoundD pptsoundd = new PptSoundD();
            PptActionTypeTimeD pptactiontypetimed = new PptActionTypeTimeD();
            PptActionOnTimeD pptactionontimed = new PptActionOnTimeD();
            PptBackgroudColorD pptbackgroudcolord = new PptBackgroudColorD();
            pptquestion.PaperType = "D";
            DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                switch (questionflag)
                {
                    case "插入幻灯片":
                        pptnewd.Move(pptquestion);
                        break;

                    case "图片大小":
                        pptpicturesized.pptpicturesize(pptquestion);
                        break;
                    case "超链接":
                        ppthyperlinkd.ppthyperlink(pptquestion);
                        break;
                    case "动画效果":
                        pptanimationeffectd.PptAnimationEffect(pptquestion);
                        break;
                    case "动画时间":
                        pptanimationeffecttimed.PptAnimationEffect(pptquestion);
                        break;
                    case "切换效果":
                        pptactiontyped.actionType(pptquestion);
                        break;
                    case "声音":
                        pptsoundd.actionType(pptquestion);
                        break;

                    case "切换时间":
                        pptactiontypetimed.actionType(pptquestion);
                        break;
                    case "自动换片时间":
                        pptactionontimed.actionTypetime(pptquestion);
                        break;
                    case "背景颜色":
                        pptbackgroudcolord.BackgroundD(pptquestion);
                        break;
                }
        
            }
        }
        #endregion

        #region  试卷E2015年12月11日16:36:54
        public void SwitchQuestionFlagE(PptQuestionEntity pptquestion)
        {
            PptAnimationEffectE pptanimationeffecte = new PptAnimationEffectE();
            PptActionTypeE pptactiontypee = new PptActionTypeE();
            PptMoveE pptmovee = new PptMoveE();
            PptThemeE pptthemee = new PptThemeE();
            PptNewE pptnewe = new PptNewE();
            PptArtWordTextE pptartwordtexte = new PptArtWordTextE();
            PptWordartE pptwordarte = new PptWordartE();
            PptWordartTypeE pptwordtypee = new PptWordartTypeE();

            pptquestion.PaperType = "E";
            DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                switch (questionflag)
                {
                    case "动画效果":
                        pptanimationeffecte.PptAnimationEffect(pptquestion);
                        break;

                    case "切换效果":
                        pptactiontypee.actionType(pptquestion);
                        break;
                    case "移动幻灯片":
                        pptmovee.Move(pptquestion);
                        break;
                    case "主题":
                        pptthemee.actionType(pptquestion);
                        break;
                    case "添加幻灯片":
                        pptnewe.Move(pptquestion);
                        break;
                    case "艺术字文本":
                        pptartwordtexte.actionType(pptquestion);
                        break;
                    case "艺术字样式":
                        pptwordarte.actionType(pptquestion);
                        break; 
                    case "艺术字形状":
                        pptwordtypee.actionType(pptquestion);
                        break;

                                      
                }

            }
        }
        #endregion

        #region  试卷F2015年12月11日16:48:11
        public void SwitchQuestionFlagF(PptQuestionEntity pptquestion)
        {
            PptAnimationEffectF pptanimationeffectf = new PptAnimationEffectF();
            PptFormatF pptformatf = new PptFormatF();
            PptMoveF pptmovef = new PptMoveF();
            PptActionTypeF pptactiontypef = new PptActionTypeF();
            PptSoundF pptsoundf = new PptSoundF();
            PptActionTypeTimeF pptactiontypetimef = new PptActionTypeTimeF();
            PptPictureSizeF pptpicturesizef = new PptPictureSizeF();
            PptAnimationEffectPictureF pptanimationeffectpicturef = new PptAnimationEffectPictureF();
            pptquestion.PaperType = "F";
            DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                switch (questionflag)
                {
                    case "动画效果":
                        pptanimationeffectf.PptAnimationEffect(pptquestion);
                        break;

                    case "幻灯片版式":
                        pptformatf.actionType(pptquestion);
                        break;
                    case "移动幻灯片":
                        pptmovef.Move(pptquestion);
                        break;
                    case "切换效果":
                        pptactiontypef.actionType(pptquestion);
                        break;
                    case "声音":
                        pptsoundf.actionType(pptquestion);
                        break;
                    case "切换时间":
                        pptactiontypetimef.actionType(pptquestion);
                        break;
                    case "图片大小":
                        pptpicturesizef.pptpicturesize(pptquestion);
                        break;

                    case "图片动画效果":
                        pptanimationeffectpicturef.PptAnimationEffect(pptquestion);
                        break;
                }

            }
        }
        #endregion

        #region  试卷G2015年12月11日16:59:05
        public void SwitchQuestionFlagG(PptQuestionEntity pptquestion)
        {
            PptNewG pptnewg = new PptNewG();
            PptHyperlinkG ppthyperlinkg = new PptHyperlinkG();
            PptActionTypeG pptactiontypeg = new PptActionTypeG();
            PptBackgroudColorG pptbackgroundcolorg = new PptBackgroudColorG();
            pptquestion.PaperType = "G";
            DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                switch (questionflag)
                {
                    case "添加幻灯片":
                        pptnewg.Move(pptquestion);
                        break;

                    case "超链接":
                        ppthyperlinkg.ppthyperlink(pptquestion);
                        break;
                    case "动画效果":
                        pptactiontypeg.actionType(pptquestion);
                        break;
                    case "背景颜色":
                        pptbackgroundcolorg.actionType(pptquestion);
                        break;
                    
                }

            }
        }
        #endregion

        #region  试卷H2015年12月11日17:17:51
        public void SwitchQuestionFlagH(PptQuestionEntity pptquestion)
        {
            PptArtWordTextH pptartwordtexth = new PptArtWordTextH();
            PptFontNameH pptfontnameh = new PptFontNameH();
            PptSizeH pptsizeh = new PptSizeH();
            PptWordartH pptwordarth = new PptWordartH();
            PptActionTypeH pptactiontypeh = new PptActionTypeH();
            PptFormatH pptformath = new PptFormatH();
            PptTextH ppttexth = new PptTextH();
            PptTextSizeH ppttextsizeh = new PptTextSizeH();
            
            PptHyperlinkH ppthyperlinkh = new PptHyperlinkH();
            PptPictureSizeH pptpicturesizeh = new PptPictureSizeH();
            PptTextWidthH ppttextwidthh = new PptTextWidthH();
            pptquestion.PaperType = "H";
            DataTable dt = pptquestionbll.LoadPptQuestion(pptquestion);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                switch (questionflag)
                {
                    case "艺术字文本":
                        pptartwordtexth.actionType(pptquestion);
                        break;

                    case "字体":
                        pptfontnameh.FontName(pptquestion);
                        break;
                    case "字号":
                        pptsizeh.size (pptquestion);
                        break;
                    case "艺术字样式":
                        pptwordarth.actionType(pptquestion);
                        break;
                    case "切换效果":
                        pptactiontypeh.actionType(pptquestion);
                        break;
                    case "幻灯片版式":
                        pptformath.actionType(pptquestion);
                        break;
                    case "文本框":
                        ppttextsizeh.TextH(pptquestion);
                        break;

                    case "文本框插入字":
                        ppttexth.actionType(pptquestion);
                        break;

                    case "超链接":
                        ppthyperlinkh.ppthyperlink(pptquestion);
                        break;

                    case "插入图片":
                        pptpicturesizeh.pptpicturesize(pptquestion);
                        break;
                    case "图片大小":
                        ppttextwidthh.PptTextWidth(pptquestion);
                        break;
                }

            }
        }
        #endregion
    }
}
