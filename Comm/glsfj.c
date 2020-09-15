#include <windows.h>
#include <stdio.h>
#include <malloc.h>
#include<string.h>
#include<conio.h>
#include "glsfj.h"

HANDLE	hThread;
HANDLE	chdthd[10];

long _EXPORT FAR PASCAL WinPROC(HWND,unsigned,UINT,LONG);
/*
 * WinMain
*/
int PASCAL WinMain(HANDLE hInst,HANDLE hPreInst,LPSTR lpcmd,int nStyle)
{
	WNDCLASS	wc;
	MSG			msg;
	HWND		hWnd;
	char		Name[]="GLSFJ";
	
	while (GetMessage(&msg,0,0,0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
    }
	return(msg.wParam);
}

//		WinProc

long _EXPORT FAR PASCAL WinProc(HAND hwnd,unsigned msg,WORD wParam,LONG lParam)
{

}

LRESULT WINAPI mix(int gl)
{

  int ljmn,ljmm=0,zu,ljml,ljmll,j,jt,msdn,ltt,ljbd,ljsj1,i,ecl=0;
  float ljmnl,zjnl,ljmjw,ljmwk,ljmnk,meq1;
  long tt=0,lt,tnl;
  time_t ljmst_1,ljmtfs_1,ljmst_2,ljmtfs_2,ljmth1,ljmth2,ljmz_k;

pfnl=0;pfhlsj=0;
if(COUNT==0)
jhcount=99;
else
jhcount=0;
for (i=0x150;i<0x158;i++) output(0x0000,i);
while(1)
{

 do
 {
   run_f=read_run(run.flag)
   if(run_f==1){
                mrun=read_produce(produce_do);
               }	   	
   sleep(1100);
 }while((run_f !=1)||(mrun !=1));


 /*  日程序开始 */
					
 mrun=1;
 ljmwk=(jcc[4][1]-jcc[4][0])/(jdd[4][1]-jdd[4][0]);/* 温度比例系数 */
 ljmnk=(jcc[5][1]-jcc[5][0])/(jdd[5][1]-jdd[5][0]);/* 能量比例系数 */


  ljmn=1;                               /* 车数记载 */
  msdn=0;                               /* 手动车数 */
  l_n=0;
			
 step1:;
 do
 {
  if (run_f==0) break;
   ljmm=0;mst=0;pfzu=0;zjnl=0;nss=ljmn;
   por15=por15&lthr;sleep(330);	       // 输入密炼机故障信息
   if ((por15&thr)==thr)
        { mixerror(28);ddcl();}		     
   por18=por18&lsix;sleep(330);       // 检查后加料器状态
/* if ((por18&six)!=six)
     { mixerror(49);ddcl();} */	
   por14=por14&lsix;sleep(440);      // 密炼机手动/自动信息
						
stop10:;
    while((por14&six)!=six)
    {
      // 手动 
	por16=por16&ltwo;
	por16=por16&lfou;
	por17=por17&lfiv;
	por14=por14&lsix;sleep(1100);
	if((por16&two)==0) yl=0; else yl=1;
	if((por16&fou)==0) th1=0; else th1=1;
	if((por17&fiv)==0) gl=0; else gl=1;
	if((por14&six)==six) goto step1;
    }

   if (ljmn>ns) break;
      plmg();                      // 关排料门
      if(jh77==1){;}
	  jlmg();                      // 关加料门
      if(jh77==1){;}
	  syl(0);                      // 上顶栓压力0.5公斤 提上顶栓
      por14=por14&lthr;
 while ((gl != 1)||(th1 != 1)||(yl != 1))
 {                                 // 物料准备好检测

   sleep(1100);
 }
 
   ljmjw=ljmwk*(jw-jdd[4][0]);mwd=ljmjw;
   por14=por14&lfiv;sleep(220);
   j=1;
 while((por14&fiv)!=fiv)          // 检查密炼机是否就绪
  {
   sleep(1100);
   j++;
   if(j>18) break;
  }
    if(j>=18) {mixerror(47);ddcl();}

  por15=por15&lone;sleep(110);    // 查密炼机转速
  if (jhmn[ljmm][7]!=0) {
   mzs=((por15&one)==one)?20:40;
   if (mzs!=jhmn[ljmm][7])
   {
     por52=por52|thr;
     output(por52,0x152);
     mixerror(40);                // 密炼机转速错
     ddcl();
   }}
     if (ljmjw<pfdw) mixerror(35);

 step4:
 do
 {
  if (run_f==0) break;

   if (jhmn[ljmm][12] > 0)
   {
     syl(0);                     //上顶栓压力0.5公斤 提上顶栓
     por14=por14&ltwo;
     if(jh77==1){;}
	 jlmk();                     // 开加料门
     por14=por14&lfou;
   if((por17&one)==one)  {       // 投胶料
	 por51=por51|eit;
     output(por51,0x151);
     sleep(jhmn[ljmm][12]*1100);
     por17=por17&lfiv;sleep(220);
    while((por17&fiv)==fiv) {output(por51,0x151);por17=por17&lfiv;sleep(2200);}
    }
   else   {
     por15=por15&ltwo;sleep(220);
     while((por15&two)!=two) sleep(220);
     }
   if(jh77==1){;}
	por51=por51&leit;
    output(por51,0x151);
    sleep(660);gl=0;
    if(jh77==1) {;}
    jlmg();                        // 关投料门
    por14=por14&lthr;
    if(ljmm==0) time(&ljmz_k);
    sleep(880);
    syl(jhmn[ljmm][8]);           // 输出上顶栓压力
   }

   if(jhmn[ljmm][10]>0)           // 投碳黑
   {
    if(th2==1)
    {if(ecl==0) {while(th==0) sleep(880);}
     ecl=ecl+1;
    }
	th3=1;sleep(55);
    por14=por14&lone;
    syl(0);                      // 压力0.5公斤 提上顶栓
  // zu=(int)(jhmn[ljmm][10]/2);
    zu=(int)(jhmn[ljmm][10]);
    por14=por14&ltwo;
    if(jh77==1) {;}
    if((por19&1)==1)  {
 // for(i=0;i<2;i++) {
	 time(&ljmth1);
	 por51=por51|six;
	 output(por51,0x151);
	if(jh77==1){;}
	do{sleep(990);time(&ljmth2);
	  }while(difftime(ljmth2,ljmth1)<zu);
 // }
	 por16=por16&lfou;sleep(550);
      if((por16&fou)==fou) mixerror(50);
	 while((por16&fou)==fou) { por16=por16&lfou;
       sleep(1100);}
    }
   else  {
     por15=por15&ltwo;sleep(220);
     while((por15&two)!=two) sleep(220);
     }
	 por51=por51&lsix;
	 output(por51,0x151);
     if(jh77==1){;}
	 th3=0;sleep(110);
     sleep(550);th1=0;
     if(th2==1) {if(ecl==2){ecl=0;th=0;}} else th=0;
     syl(jhmn[ljmm][8]);         // 输出上顶栓压力
   }

   if (jhmn[ljmm][11] > 0)       // 投油料
   {
    if(jh77==1) {;}
    if (jhmn[ljmm][11] == 1)
    {
     syl(0);
     jlmk();
     ddcl();
     jlmg();
    }
   else
   {
    if(pflw2==1)  syl(0);       // 压力0.5公斤 提上顶栓
    if((por1a&1)==1) {
	por51=por51|siv;
    output(por51,0x151);
    if(jh77==1){;}
    sleep(jhmn[ljmm][11]*1100);
    por51=por51&lsiv;
    output(por51,0x151);
    p56d3_1();sleep(3000);      // 吹管路
    p56d3_0();
   if(jh77==1){;}
    }
   else
    {
     por15=por15&ltwo;sleep(220);
     while((por15&two)!=two) sleep(220);
    }
    sleep(220);yl=0;
    syl(jhmn[ljmm][8]);         // 输出上顶栓压力
   }
   }

   time(&ljmtfs_1);
   tt=difftime(ljmtfs_1,ljmst_1);

	if (jhmn[ljmm][2]==0 && jhmn[ljmm][4]==0 && jhmn[ljmm][6]==0) ljbd=1;
	if (jhmn[ljmm][2]==0 && jhmn[ljmm][4]==0 && jhmn[ljmm][6]==1) ljbd=2;
	if (jhmn[ljmm][2]==0 && jhmn[ljmm][4]==1 && jhmn[ljmm][6]==0) ljbd=3;
	if (jhmn[ljmm][2]==0 && jhmn[ljmm][4]==1 && jhmn[ljmm][6]==1) ljbd=4;
	if (jhmn[ljmm][2]==1 && jhmn[ljmm][4]==0 && jhmn[ljmm][6]==0) ljbd=5;
	if (jhmn[ljmm][2]==1 && jhmn[ljmm][4]==0 && jhmn[ljmm][6]==1) ljbd=6;
	if (jhmn[ljmm][2]==1 && jhmn[ljmm][4]==1 && jhmn[ljmm][6]==0) ljbd=7;
	if (jhmn[ljmm][2]==1 && jhmn[ljmm][4]==1 && jhmn[ljmm][6]==1) ljbd=8;

   ljml=0;lt=0;ljmll=0;ltt=0;mds=ljmm+1;mdt=0;mzu=0;
   syl(jhmn[ljmm][8]);                      // 输出上顶栓压力
   time(&ljmst_1);                          // 开始记时
   ljmtfs_2=ljmst_1;
   meq=0.0;                                 // 输入开始能量
/* step3: */
  do
   {
   if (run_f==0) break;
   por14=por14&lsix;sleep(330);
   if((por14&six)!=six) ddcl();
    /* jczt(); */
    sleep(220);
    time(&ljmtfs_1);
    tt=difftime(ljmtfs_1,ljmst_1);           // 本段时间到否
    tnl=difftime(ljmtfs_1,ljmtfs_2);         // 能量时间
    mst=difftime(ljmtfs_1,ljmz_k);/* mst=mst+tt-mdt; */
    mdt=tt;mzu=(int)(mdt*mzs/60);pfzu=(int)(mst*mzs/60);
					                         // 累记总时间及本车总转数
	ljmjw=ljmwk*(jw-jdd[4][0]);mwd=ljmjw;
	if (ljmjw>pfgw)                         // 超温报警
       { mixerror(36);ljmll=1;ljml=1;output(por50|eit,0x150);}
	if (ljmll==1) break;
	if (ljmjw<pfdw) mixerror(35);           // 胶温过低
	if (tnl>0) {ljmtfs_2=ljmtfs_1;ljmnl=ljmnk*(nl-jdd[5][0])*tnl/3600.0;meq=meq+ljmnl;zjnl=zjnl+ljmnl;meq1=meq*10.0;}
 if (ljsj1==1)
 {
   if (ltt==0) {time(&ljmst_2);ltt=1;}
   time(&ljmtfs_1);lt=difftime(ljmtfs_1,ljmst_2);
   if (lt>pfsj) {mixerror(41);ljml=1;lt=0;ltt=0;ljsj1=0;} /* 连接没有成功 */
  }
    sleep(550);
 if (ljml==1) break;
 switch(ljbd)
	{
	  case 1:
	   if (tt>=jhmn[ljmm][1]||ljmjw>=jhmn[ljmm][3]||meq1>=jhmn[ljmm][5]) {ljml=1;break;}
	   break;
	  case 2:
	   if ((tt>=jhmn[ljmm][1]||ljmjw>=jhmn[ljmm][3])&&meq1>=jhmn[ljmm][5]) {ljml=1;break;}
	   if (tt>=jhmn[ljmm][1]||ljmjw>=jhmn[ljmm][3])
	    { ljsj1=1;
	      if (tt>=jhmn[ljmm][1]) mixerror(37);
	      if (ljmjw>=jhmn[ljmm][3]+5) mixerror(39);
        }
	    if (meq1>=jhmn[ljmm][5])
	     { ljsj1=1;
	       if (ljml=1) break;
	       mixerror(38);
         }
	   break;
	  case 3:
	   if ((tt>=jhmn[ljmm][1]||meq1>=jhmn[ljmm][5])&&ljmjw>=jhmn[ljmm][3]) {ljml=1;break;}
	   if (tt>=jhmn[ljmm][1]||meq1>=jhmn[ljmm][5])
	    { ljsj1=1;
	      if (tt>=jhmn[ljmm][1]) mixerror(37);
	      if (meq1>=jhmn[ljmm][5]) mixerror(38);
        }
	    if (ljmjw>=jhmn[ljmm][3])
	     { ljsj1=1;
	       if (ljml=1) break;
           if (ljmjw>=jhmn[ljmm][3]+5) mixerror(39);
	     }
	   break;
	  case 4:
	   if (tt>=jhmn[ljmm][1]||(meq1>=jhmn[ljmm][5]&&ljmjw>=jhmn[ljmm][3])) {ljml=1;break;}
	   if (tt>=jhmn[ljmm][1])
	    { ljsj1=1;
	      mixerror(37);
	    }
	    if (ljmjw>=jhmn[ljmm][3]&&meq1>=jhmn[ljmm][5]) ljsj1=1;
	      if (meq1>=jhmn[ljmm][5]) mixerror(38);
	      if (ljmjw>=jhmn[ljmm][3]+5) mixerror(39);
	    break;
	  case 5:
	   if (tt>=jhmn[ljmm][1]&&(meq1>=jhmn[ljmm][5]||ljmjw>=jhmn[ljmm][3])) {ljml=1;break;}
	   if (ljmjw>=jhmn[ljmm][3]||meq1>=jhmn[ljmm][5])
	    { ljsj1=1;
	      if (ljmjw>=jhmn[ljmm][3]+5) mixerror(39);
	      if (meq1>=jhmn[ljmm][5]) mixerror(38);
	    }
	    if (tt>=jhmn[ljmm][1])
	     { ljsj1=1;
	       if (ljml=1) break;
	       mixerror(37);
	     }
	   break;
	  case 6:
	   if (ljmjw>=jhmn[ljmm][3]||(meq1>=jhmn[ljmm][5]&&tt>=jhmn[ljmm][1])) {ljml=1;break;}
	   if (ljmjw>=jhmn[ljmm][3])
	    { ljsj1=1;
	      if (ljmjw>=jhmn[ljmm][3]+5) mixerror(39);
	    }
	    if (tt>=jhmn[ljmm][1]&&meq1>=jhmn[ljmm][5]) ljsj1=1;
	      if (meq1>=jhmn[ljmm][5]) mixerror(38);
	      if (tt>=jhmn[ljmm][1]) mixerror(37);
	    break;
	  case 7:
	   if (meq1>=jhmn[ljmm][5]||(ljmjw>=jhmn[ljmm][3]&&tt>=jhmn[ljmm][1])) {ljml=1;break;}
	   if (meq1>=jhmn[ljmm][5])
	    { ljsj1=1;
	      mixerror(38);
	    }
	    if (tt>=jhmn[ljmm][1]&&ljmjw>=jhmn[ljmm][3]) ljsj1=1;
	      if (ljmjw>=jhmn[ljmm][3]+5) mixerror(39);
	      if (tt>=jhmn[ljmm][1]) mixerror(37);
	    break;
	  case 8:
	   if (meq1>=jhmn[ljmm][5]&&(ljmjw>=jhmn[ljmm][3]&&tt>=jhmn[ljmm][1])) {ljml=1;break;}
	   if (tt>=jhmn[ljmm][1])
	    { ljsj1=1;
	      mixerror(37);
	    }
	   if (ljmjw>=jhmn[ljmm][3])
	    { ljsj1=1;
	      if (ljmjw>=jhmn[ljmm][3]+5) mixerror(39);
	    }
	   if (meq1>=jhmn[ljmm][5])
	    { ljsj1=1;
	      mixerror(38);
	    }
	    break;
	}

   sleep(550);
   }while (ljml != 1);          /* 对应 step3: */

  if (run_f==0) break;
  d_ml[ljmm][0]=mdt;d_ml[ljmm][1]=mwd;d_ml[ljmm][2]=(int)(meq);
   if (ljmll == 1) goto step7;  /*  如胶温超高卸料开始下一车  */

  ljmm=ljmm+1;
  if (ljmm == pfmn-1)           // * 准备排胶 *
  {
   por14=por14&lsiv;
   por52=por52|two;
   output(por52,0x152);
   sleep(220);
  }
  }while(ljmm < pfmn);              /* 对应 step4: */
  if (run_f==0) break;

  if((por14&siv)!=siv) {mixerror(48);syl(0);}
  while((por14&siv)!=siv)
   {sleep(990);}                    /* 下辅机未就绪 */
  /*syl(jhmn[ljmm][8]);             输出上顶栓压力  */
    por52=por52&ltwo;               /* 准备排胶信号清零 */
    output(por52,0x152);

step7:;
  pfplwd=ljmwk*(jw-jdd[4][0]);      /* 排料时胶温 */
  tnl=difftime(ljmtfs_1,ljmtfs_2);ljmnl=ljmnk*(nl-jdd[5][0])*tnl/3600.0;
  meq=meq+ljmnl;pfnl=zjnl+ljmnl;    /* 本车总耗能量 */
  pfhlsj=mst;mst=pfhlsj;            /* 本车炼胶总时间  */
  pfzu=(int)(mst*mzs/60);           /* 本车总转数 */
/*if (tt>pfzsj) mixerror(42);*/
  if(pflw1==1) syl(0);              /* 压力0.5公斤 提上顶栓 */
  if(jh77==1) {;}
  plmk();                           // 下顶栓开卸料!
  por15=por15&lsiv;
  if(pflw1 != 1) syl(0);            /* 压力0.5公斤 提上顶栓 */
  if(d_dy=='Y') {zss=ljmn;jhprint=14;}
   sleep(pfpl*1100);                /* 排料开始延时  */
   output(por50&leit,0x150);        /* 超温报警信号清零 */
   if(l_n == 0) l_n=1; else l_n=0;
   ljmn=ljmn+1;
   if (run_f==0) break;
   if((por14&six)!=six) goto stop10; 
  }while (ljmn<=ns);               /* 对应 step1: */
 
	por50=por50&lthr|fiv;
	output(por50,0x150);
/*if((run_f==1)&&(ljmrun==1)) renh();*/
  mrun=0;mrun1=0;               /* 通知其它任务一个配方运行完毕 */
  sleep(880);

}

}

/*   输出上顶栓压力   */
syl(int x)
{
	int j;

   switch(x)
   {
    case 0:                    // 上顶栓压力 0
	por14=por14&lone;
	j=0;
     por52=por52&0x000f|six;output(por52,0x152);
     por50=por50&lsiv|six;output(por50,0x150);myl=0;
      while((por14&one)!=one)    
      {    
	  sleep(880);
       j++;if(j>8) break;
      };
     if(jh77==1){_setcolor(14);_rectangle(3,379,225,402,275);}
     break;
    case 1:                   // 上顶栓压力 1
	por14=por14&ltwo;
    j=0;
	do{
     por50=por50&lsix|siv;output(por50,0x150);
     por52=por52&0x000f|eit;output(por52,0x152);myl=1;
	  sleep(880);
       j++;if(j>4) break;
      }while((por14&two)!=two);
     if(jh77==1){;}
     break;
    case 2:                  // 上顶栓压力 2
	por14=por14&ltwo;
	 j=0;
	 do{
     por52=por52&0x000f|siv;output(por52,0x152);myl=2;
     por50=por50&lsix|siv;output(por50,0x150);
	  sleep(880);
       j++;if(j>4) break;
      }while((por14&two)!=two);
     if(jh77==1){;}
     break;
    case 3:                 // 上顶栓压力 3 
	por14=por14&ltwo;
	 j=0;
	 do{
     por52=por52&0x000f|six;output(por52,0x152);myl=3;
     por50=por50&lsix|siv;output(por50,0x150);
	  sleep(880);
       j++;if(j>4) break;
      }while((por14&two)!=two);
     if(jh77==1){;}
     break;
    case 4:                // 上顶栓压力 4
	por14=por14&ltwo;
	 j=0;
	 do{
     por52=por52&0x000f|fiv;output(por52,0x152);myl=4;
     por50=por50&lsix|siv;output(por50,0x150);
	  sleep(880);
       j++;if(j>4) break;
      }while((por14&two)!=two);
     if(jh77==1){;}
     break;
   }
 }

jlmk()
{
 int j;
	por14=por14&lthr;
 j=1;
 while((por14&thr)!=thr)
  {
      por51=por51&lthr|fou;
      output(por51,0x151);
   sleep(1100);
   j++;
   if(j>18) break;
  }
   if(j>=18) {mixerror(33);ddcl();}
	mjlm=1;
}
jlmg()
{
 int j;
      por14=por14&lfou;
 j=1;
 while((por14&fou)!=fou)
  {
      por51=por51&lfou|thr;
      output(por51,0x151);
   sleep(1100);
   j++;
   if(j>18) break;
  }
    if(j>=18) {mixerror(34);ddcl();}
	mjlm=0;
}
plmk()
{
 int j;
   por15=por15&leit;
 j=1;
 while((por15&eit)!=eit)
  {
   por50=por50&lfiv|thr;
   output(por50,0x150);
   sleep(880);
   j++;
   if(j>20) break;
  }
      if(j>=20) {mixerror(32);ddcl();}
	mxlm=1;                             /* 排料门开 */
}
plmg()
{
 int j;
      por15=por15&lsiv;
 j=1;
 while((por15&siv)!=siv)
  {
      por50=por50&lthr|fiv;
      output(por50,0x150);
   sleep(1100);
   j++;
   if(j>18) break;
  }
   if(j>=18) {mixerror(31);ddcl();}
	mxlm=0;                             /* 排料门关 */
}

ddcl()                                  // 等待继续运行信号
{
 int zu;
    por15=por15&ltwo;zu=0;
    sleep(2200);/*mixerror(46);*/
    do
    {
     zu=por15&two;
     sleep(880);
    }while(zu!=two);
     por14=por14&lsix;
}

/* jczt()
 {
  while((por16&one)==one)
   {mixerror(46);ddcl();}
    por16=por16&lone;
 } */

mixerror(int x)
{
		#if COUNT==0
			locat1(16,54);
			switch(x)
			{
			  case 28:
				cor(0x8e,"     密炼机故障     ");
				break;
			   case 29:
				cor(0x8e,"   上顶栓提起故障   ");
			   case 30:
				cor(0x8e,"   上顶栓下降故障   ");
			   case 31:
				cor(0x8e,"    排料门关故障    ");
				break;
			   case 32:
				cor(0x8e,"    排料门开故障    ");
				break;
			   case 33:
				cor(0x8e,"    进料门开故障    ");
				break;
			   case 34:
				cor(0x8e,"    进料关开故障    ");
				break;
			   case 35:
				cor(0x8e,"温度低于配方最低温度");
				break;
			   case 36:
				cor(0x8e,"   密炼机超温报警   ");
				break;
			   case 37:
				cor(0x8e," 时间超过配方规定值 ");
				break;
			   case 38:
				cor(0x8e," 能量超过配方规定值 ");
				break;
			   case 39:
				cor(0x8e," 温度超过配方规定值 ");
				break;
			   case 40:
				cor(0x8e,"    密炼机转速错    ");
				break;
			   case 41:
				cor(0x8e,"  过程连接没有成功  ");
				break;
			   case 42:
				cor(0x8e,"   混炼总时间超过   ");
				break;
			   case 43:
				cor(0x8e,"    进料时间超过    ");
				break;
			   case 44:
				cor(0x8e," 粉中间斗排料不完全 ");
				break;
			   case 45:
				cor(0x8e,"   粉秤中间斗故障   ");
				break;
			   case 46:
				cor(0x8e,"  等待继续运行信号  ");
				break;
			   case 47:
				cor(0x8e,"    密炼机未就绪    ");
				break;
			   case 48:
				cor(0x8e,"    下辅机未就绪    ");
				break;
			   case 49:
				cor(0x8e,"   后加料器状态错   ");
				break;
			   case 50:
				cor(0x8e,"   炭黑秤中间斗故障 ");
				break;
			}
		#else
			 jherror(x);
		#endif
}

LRESULT WINAPI rub(int gl)       /* ******** 生胶秤 ******** */
/* ------------------------------------------------------------- */

{
int i,j,k,n,m,nn,l,lc,ik,mgl1,mgl2,d_j=0;
float sjzc,mb;
  /* xji4=0;    秤皮带光电信号  */
  /* xji5=1;    投料皮带光电信号  */
 scsh(0xe3);    /* 通讯口初始化,终端清屏  */
outc(0x1b);
outc(0x5b);
outc('#');
outc('Y');
while(1)
{
	do
	{
	 sleep(1100);gl=1;
	}while((mrun != 1)||(sj1_sum+sj2_sum == 0));
	nn=1;gl=0;mgl1=0;          /*   ns=车数   */
	m=pfs3n;  /*  pfs3n 为胶的种数  */
	while(s3[mgl1].jl==1) mgl1++;
	/*mgl2=pfs3n-mgl1;*/
	for(i=0;i<m;i++)
	s3[i].sz=0;       /* 实重数组冲零  */

  /* ---------------------------------------------------------- */
while(nn<=ns)       /*  对车数开始循环  */
{
por17=por17&0x40;
sleep(55);
m=pfs3n;
do
	{
	/* s3l=(sj-jdd[2][0])*LJF; */ ljmsj_1();sjzc=s3l;
	 sleep(550);
	 ljmsj_1();
	}while((s3l-sjzc)>0.4);
	/* }while(fabs(s3l-(sj-jdd[2][0])*LJF)>1.0);  皮重  */
	/* ------------------------------------ */
	for(i=0;i<pfs3n;i++)    /*  对胶种数循环开始  */
	 {
pp135();
np=itoa((ns-nn+1),nnn,10);
if((ns-nn+1)<10)
pp4(6,1,2);
if((ns-nn+1)>9&&(ns-nn+1)<100)
pp4(6,1,1);
if((ns-nn+1)>99)
pp4(6,1,0);
for(l=0;l<3;l++)
outc(nnn[l]);           /* 送车数到终端 */
pp4(2,0,3);
for(l=0;l<16;l++)
	{
	sleep(55);
    outc(s3[i].mi[l]);  /* 当前胶种送终端 */
	}
 mb=s3l;
 s3w3=s3[i].zl+s3l;    /* s3w3=当前胶种重量+秤上的其他胶种重量 */
 s3w6=s3[i].gc;        /* s3w6= 当前胶种的配方公差  */
 sleep(110);
 pp2(s3[i].zl);        /* 终端显示当前胶种的配方重量  */
 if (m==pfs3n)
   {
	if ((por17&one)==one)
	{                 /* 手动时秤带不启动 */
	 por53=por53|eit;
	 output(por53,0x153);   /*  起动秤带慢进  */
	 lw2=0;             /* 秤带慢速起动信号  */
	 if(jh77==1){;}
	}
   }
 s3run=2;    /* 秤运行 */
 por17=por17&ltwo;
 por17=por17&lone;sleep(220);
if((por17&one)==one)
{                                       /* 自动称量   */
 s3[i].zt=4;
 do                   /*  开始称胶  */
  {
	/*stop();*/
	sleep(55);
   do                 /* 对秤的两次称量作比较  */
   {
	s3l1=0.0;
	ljmsj_1();sjzc=s3l;     /* s3l 当前胶秤实际重量 */
	s3w1=s3w3-s3l;
	/*stop();*/
	sleep(440);
    pp2(s3w1);       /* 终端显示 +,- 量值, +值表示加料, -值表减料  */
    xji4=por17&4;    /*  读秤皮带光电信号, 1 停止, 0 进 !  */
    por17=por17&lthr;
    if ((xji4==4)||(s3w1<28.0))
     {
      por53=por53&leit;
      output(por53,0x153);  /*  停慢速带, 停供胶机 */
	  lw2=1;
	  if(jh77==1){;}
     }
     ljmsj_1();
   }while(fabs(s3l-sjzc)>1.0);  /* 先后两次称量>1.0 重读 */
	if(fabs(s3w1)<s3w6)
    {
	 ik=1;
	 por17=por17&ltwo;sleep(55);
	 while((ik<3)&&((por17&two)!=two))
      {
	   /*s3l=(sj-jdd[2][0])*LJF;*/ ljmsj_1();
	   s3w1=s3w3-s3l;
	   pp2(s3w1);
	   sleep(1800);ik=ik+1;
	  }
	}
	por17=por17&lone;sleep(220);
	if((por17&one)!=one) goto done2;
done1:;
   }while(fabs(s3w1)>s3w6&&((por17&two)!=2)); /* 小于公差或按接收钮则称胶结束  */
	if(m==pfs3n)
	{
	 por53=por53&leit;
	 output(por53,0x153);
	 }
}
	else             /*  手动时执行下段程序  */
 {
  s3[i].zt=3;
  por17=por17&ltwo;sleep(110);
  do
   {                       /* 手动称量  */
    s3l1=0.0;
    ljmsj_1();
    s3w1=s3w3-s3l;
    sleep(110);pp2(s3w1);
	por17=por17&lone;sleep(220);
	if((por17&one)==one) goto done1;
 done2:;
   }while((por17&two)!=two);por17=por17&ltwo;
  }       /*  手动时,每称完一种胶要按重量接收键   */
  m=m-1;
  d_sz[d_j][2][i]=s3l-mb;
  sleep(1000);
  s3[i].sz=s3[i].sz+s3[i].zl-s3w1;
  if((i==mgl1-1)||(i==pfs3n-1))
	{
	sleep(1000);
    if((gl==1)||(por17&fiv)==fiv) pp136(1); /* 终端显示: 等待卸胶 ! */
	/* ------------------------------------ */
	sleep(330);
    s3run=1;/*  送秤结束信号  */
	if((por17&one)!=one)  /* 手动时执行下面程序 */
	{
    do
     {
	  por17=por17&lfiv;
      sleep(550);
	 }while((gl==1)||(por17&fiv)==fiv);

    pp136(2);
	por17=por17&lfiv;sleep(165);
	while((por17&fiv)!=fiv)
	sleep(440);
   }   /*  手动时,由操作者启动把秤上的胶料御到投料带上  */
	else
   {
	while((gl==1)||(por17&fiv)==fiv)
	{
	por17=por17&lfiv;
	sleep(550);
	};
    sleep(1100);

    pp136(2);
	por53=por53|six;
	output(por53,0x153); /* 快速起动秤皮带,把秤上的料御到投料带 */
	if(jh77==1){;}
	sleep(1100);
	xxc=0;lw3=0;
	por17=por17&lfiv;
	while((por17&fiv)!=fiv)
	{
	/*stop();*/
	sleep(550);
	};
	por53=por53&lsix;
	output(por53,0x153);   /* 停高速, 投料带 */
	lw3=1;
	if(jh77==1){;}
    pp136(3);     /* 清除终端屏幕, 显示" 等  待 ! "  */
	sleep(1100);
   }     /* else end  */
	gl=1;
  }                    /* i=mgl1 || i=pfs3n 结束 */
 }                     /*  胶种数循环结束  */

	sleep(770);
	if(d_j == 0) d_j=1; else d_j=0;
	nn++;
}       /*  车数循环完毕  */
   /* -------------------------------------------------------- */
pp136(3);             /* 清除终端屏幕, 显示" 等  待 ! "  */
scsh(0xe3);           /* 通讯口初始化  */
/*#if COUNT==0
	{
	lclcls(4,10,28,48);     清生胶秤窗口
locat1(4,30);
cor(0x1a," 材料消耗统计表: ");
locat1(6,26);
cor(0x0a,"  编号: 标准量: 实耗量: ");
	}
#endif
for(i=0;i<pfs3n;i++)
{
#if COUNT==0
	{
locat1(7+i,27);
printf("%s%7.2f %7.2f",s3[i].bh,s3[i].zl*ns,s3[i].sz);
	}
#endif
} */
do
{
 sleep(880);
}while(mrun==1);
}

}            /* ******* 生胶秤结束 ******* */
 
ljmsj_1()
{
 int k1;
 sleep(220);
 k1=atoi(ljchc);
 if(ufhc==1) k1=-k1;
 s3l=k1/10.0;
}
  
/* ------------------------------------------------------------- */

/* ------------------------------------------------------------- */

LRESULT WINAPI oil(int gl)       /* ******** 油料秤 ******** */
{
float w1,w2,w3;
float weight,quick,point,error;
int n,m,nz,i,k11,sum,ylor,d_y=0,ysl=0;
time_t start,finish,t1,t2;
go:;    por56=0;por57=0;
	do
	{
	    yl=1;
	    sleep(1100);
	}while(ljmrun!=1||pfs2n==0);
	yl=0;
/*  for(i=0;i<pfs2n;i++) {
	for(n=i+1;n<pfs2n;n++)
	if(s2[n].zl>s2[i].zl) { w1=s2[n].zl;s2[n].zl=s2[i].zl;
	   s2[i].zl=w1;  }
	}   */
	for(i=0;i<pfs2n;i++) s2[i].sz=0;
/*szzd:;*/
	n=1;yl=0;k11=0;
	yl_e1=io1(0);
	yma20();
done:;                  /* 每车料称重入口 */
	p1ad0_0();p1ad0();
if(nz==0)              // 手动
	{
/*#if COUNT==0
	   
	   locat1(10,3);
	   printf("      N=%d      ",n);
	   locat1(11,6);cor(0x8e,"              ");
	   lclcls(12,2,2,23);
	   locat1(12,4);cor(0x8e,"按重量接收键表示:");
	   locat1(13,4);cor(0x8e,"  手动时称好一车.");
	   
#endif
	   for(i=0;i<=pfs2n;i++) s2[i].zt=3;
	   do
	      {
	       p1ad1_0();
	       por16=por16&ltwo;
	       sleep(2200);p1ad0_0();p1ad0();
	       if((por1a&two)==two) {k11++;n++;}
	       if(nz != 0) goto done;
	       if((por16&two)==0) yl=0; else yl=1;
	      }while(nz==0);
*/
          sleep(2200);
	}

	m=0;

done1:;                 /* 称料入口 */
	s2[m].zt=4;
	i=s2[m].th;ylor=yad;
	weight=s2[m].zl*100.0+yad+ysl; ysl=0;
	quick=s2[m].km*100.0;
	point=s2[m].tj*100.0;
	error=s2[m].gc*100.0;
	time(&start);

	w1=0;
	p57d_1();                       /* 选 i 斗开 */
	p56d7_1();                      /* 启动油料快加阀 */
	time(&t1);sum=0;
    yma20();          /* 读A/D值 */
    while((yad+quick)<weight)
     {
	   time(&t2);
	/* while((difftime(t2,t1)>5)&&(yad<1))  {
		p56d7_0();
		sleep(1000);
		yma20(); p56d0_1();
	       yterror(YLG);ddc2();  p56d0_0();
	    } */
	    p57d_1();                         /* 选 i 斗开 */
	    p56d7_1();
	    ycon();
	    sleep(165);
	    yma20();          /* 读A/D值 */
	    if(yad>weight) yma20;
	 }

	p57d_0();               /* 选 i 斗关 */
	p56d7_0();              /* 关闭快加阀 */
	ylxg(i);
	sleep(1500);
	yma20();w1=yad-w1;

	w2=0;
	p56d6_1();                      /* 启动油料慢加阀 */
	p57d_1();                       /* 选 i 斗开 */
	time(&t1);sum=0;
    yma20();          /* 读A/D值 */
    while((yad+point)<weight)
     {
      time(&t2);
	/* while((difftime(t2,t1)>5)&&(yad<1))  {
		p56d6_0();
		sleep(1000);
		yma20(); p56d0_1();
	    yterror(YLG); ddc2();  p56d0_0();
	   }*/
	  p56d6_1();
	  p57d_1();
	  ycon();
	  sleep(55);
	  yma20();          /* 读A/D值 */
	  if(yad>weight) yma20;
	 }
	p57d_0();                       /* 选 i 斗关 */
	p56d6_0();                      /* 关闭慢加阀 */
	ylxg(i);
	sleep(1500);
	yma20();w2=yad-w2;
	ycon();
	w3=0;
	if((weight-yad)>error)        // 油料点动
	  {
	   while((weight-yad)>error){
	    w3=yad;
	    p57d_1();                   /* 选 i 斗开 */
	    sleep(220);
	    p56d6_1();
	    ylxk(i);
	    sleep((int)(jcc[1][4]*1000));
	    p57d_0();                   /* 选 i 斗关 */
	    p56d6_0();
	    ylxg(i);
	    sleep((int)jcc[1][5]*1000);
	    yma20();
	/*  while(((yad-w3)/100.0)<0)  {
	     sleep(110);
	     yma20();
	    } */
	    ycon();
	    w3=yad-w3;
	    if(yad>weight) yma20();
	  }
	 }
	sleep(55);ycon();
	d_sz[d_y][1][m]=(yad-ylor)/100.0;
	s2[m].sz=(yad-ylor)/100.0+s2[m].sz;
/*  s2[m].tj=(1.5*w2+error)/1000;
	s2[m].km=s2[m].tj+(1.5*w1)/1000;
   #if COUNT==0
	locat1(1,1);
	printf("  w1=%5.2f  w2=%5.2f  w3=%5.2f",w1/100,w2/100,w3/100);
    #endif
*/
	if(fabs(weight-yad)>error)     // 超公差
	{
     p57d0_1();p1ad4_0();p1ad2_0();
     yterror(YLJ);
     do
      {
       p1ad4();i=nz;p1ad2();
       sleep(55);
      }while((i|nz)==0);
	 p57d0_0();
	 if(nz==0)  {
	  m=0;
	  yma20();
	  s2[m].sz=s2[m].sz-yad/100.0;ysl=-yad;yl=1;
	  goto go22;
     }
	}
	m++;
	if(m<pfs2n)
	  goto done1;
	/*time(&finish);
	if(difftime(finish,start)>((int)jcc[1][3])) yterror(YLS);*/
	p16d1_0();p16d1();p1ad1_0();
   do {                      // 等待卸料
       p1ad1();sleep(880);
	   /* p16d1();
	   }while((nz!=0)||(yl==1)); */
	   }while((nz!=two)&&(yl==1));
	sleep(5500);
	p56d5_1();              /* 排油阀开,油料排下 */
	if(jh77==1){;}
   do {
	   yma20();
	   p56d5_1();              /* 排油阀开 */
	   sleep(1000);
	   p16d1();
	  }while(yad>20);
	sleep(3300);
	p56d5_0();              /* 排油阀关 */
	if(jh77==1){;}
	sleep(1500);
	yl_e1=io1(0);
	yma20();
	m=0;
go22:;
	yl=1;
	p1ad0_0();p1ad0();
	if(d_y == 0) d_y=1; else d_y=0;
	n++;
	if(ns>=n)
	  goto done;
	for(i=0;i<pfs2n;i++) {
	s2[i].sz=k11*s2[i].zl+s2[i].sz; }
/* if COUNT==0
	lclcls(8,6,1,24);locat1(8,6);
	cor(0x8e,"油料秤完毕  ");
	locat1(9,2);cor(0x0a,"斗号  标准量  实耗量");
	for(i=0;i<pfs2n;i++) { locat1(9+i,2);n=s2[i].th;
	printf("%d 斗 %7.2f %7.2f",n,ns*s2[i].zl,s2[i].sz); }
	
#endif */
	do
	{
	  sleep(880);
	}while(mrun==1);
	/* p5ad4_0(); */ s2run=1;
	goto go;
}

yma22()
{
float x,y;
	yad=io1(0)-yl_e1;
	/*s2l=yad/100.0;*/
}

yma20() /* A/D值摸拟子程序 */
{
float x;int i;
	i=atoi(ljchb);
	if(ufhb==1) i=-i;
	yad=i-yl_e1;
	/*s2l=yad/100.0;*/
}

ylxk(int i)
{
 switch(i)
 {
  case 1:
		if(jh77==1){_setcolor(4);_ellipse(3,135,177,151,193);}
		break;
  case 2:
		if(jh77==1){_setcolor(4);_ellipse(3,169,177,185,193);}
		break;
  case 3:
		if(jh77==1){_setcolor(4);_ellipse(3,203,177,219,193);}
		break;
  case 4:
		/*if(jh77==1){_setcolor(4);_ellipse(3,135,177,151,193);}*/
		break;
 }
}

ylxg(int i)
{
 switch(i)
 {
  case 1:
		if(jh77==1){_setcolor(15);_ellipse(3,135,177,151,193);}
		break;
  case 2:
		if(jh77==1){_setcolor(15);_ellipse(3,169,177,185,193);}
		break;
  case 3:
		if(jh77==1){_setcolor(15);_ellipse(3,203,177,219,193);}
		break;
  case 4:
		/*if(jh77==1){_setcolor(15);_ellipse(3,135,177,151,193);}*/
		break;
 }
}

/* ------------------------------------------------------------- */

LRESULT WINAPI carb(int gl)       /* ******** 炭黑秤 ******** */
{
float w1,w2,w3,p[6];
float weight,quick,point,error;
int n,m,nz,i,k22,ms1,ms2,jcc04,sum,thor,d_t=0,thsl=0;
time_t start,finish,t1,t2;
go:;    por54=0;por55=0;p53d2_0();p53d1_0();ms1=0;
	do
	{
	/*  p56d7_1();  */
	    th=1;th1=1;th2=1;
	 sleep(1100);
	}while(mrun!=1||pfs1n==0);
	for(i=0;i<pfs1n;i++) s1[i].sz=0;
/*szzd:;*/
	th=0;th1=0;th2=0;
	while(s1[ms1].jl==1) ms1++;
	ms2=pfs1n-ms1;
	if(ms2!=0) th2=1;
	for(i=0;i<ms1;i++) {
	for(n=i+1;n<ms1;n++)
	if(s1[n].zl>s1[i].zl) { 
	   w1=s1[n].zl;s1[n].zl=s1[i].zl;s1[i].zl=w1;
	   w1=s1[n].km;s1[n].km=s1[i].km;s1[i].km=w1;
	   w1=s1[n].tj;s1[n].tj=s1[i].tj;s1[i].tj=w1;
	   w1=s1[n].gc;s1[n].gc=s1[i].gc;s1[i].gc=w1;
	   m=s1[n].th;s1[n].th=s1[i].th;s1[i].th=m;
	  }
	 }
     n=1;k22=0;
	tw_e1=io1(1);
	k=1;
	ma20();
done:;                  /* 每车料称重入口 */
	p19d0_0();p19d0();
	if(nz==0)           /*  手动检测,按重量接收键表示手动称好一车.  */
	  {
	for(i=0;i<=pfs1n;i++) s1[i].zt=3;
	do
	   {
	     p19d4_0();
	     p16d3_0();
	     sleep(2200);p19d0_0();p19d0();
	     if((por19&fiv)==fiv) {k22++;n++;}
	     if(nz != 0) goto done;
	     if((por16&fou)==fou) {th=1;th1=1;} else th1=0;
	    }while(nz==0);
	  /* if((por16&fou)==fou) th1=1; else th1=0; */
	}

	/* p56d7_0();p56d6_1(); */ s1run=2;     /* 运行灯开 */
	m=0;

done1:;                 /* 称料入口 */
	s1[m].zt=4;
	i=s1[m].th;thor=ad;
	weight=s1[m].zl*10.0+ad+thsl; thsl=0;
	quick=s1[m].km*10.0;
	point=s1[m].tj*10.0;
	error=s1[m].gc*10.0;
	p18d0_0();p18d0();
	if(nz!=1) {yterror(THD);/*p54d5_1();*/ ddc1();/*p54d5_0();*/}
	time(&start);
	
	w1=0;
	if(m==1) {sleep(1100);p55d_7();p53d2_1();sleep(3300);p53d2_0();p55d_0();}  /*7斗开 */
   if(s1[m].zl>s1[m].km)  {
	p55d_1();               /* 选 i 斗开 */
	sleep(1000);
	p53d2_1();              /* 启动炭黑快加阀 */
	//  if(jh77==1) {locat1(20,33);cor(0x1b,"粉料快加");tlxk(i);}

	time(&t1);sum=0;
	while((ad+quick)<weight)
	   {
	    while(th3==1) { p53d2_0();sleep(550); }
	      ma20();           /* 读A/D值 */
	      time(&t2);
	  /*    while((difftime(t2,t1)>6)&&((ad)<2))     {
		   p53d2_0();
		   sleep(1000);
		ma20();   p54d5_1();
	      yterror(THD);ddc1();   p54d5_0();
	      } */
	      p53d2_1();
	   p55d_1();
	      con();
	      sleep(165);
	  if((ad)>weight) ma20();
	   }
	w1=ad;
	p53d2_0();                      /* 关闭快加阀 */
	tlxg(i);
	sleep(1000); 
	p55d_0();                       /* 选7,i号斗关 */
	sleep(1500);
	ma20();w1=ad-w1;
  }
	w2=0;
   if(s1[m].zl>s1[m].tj) {
	p55d_1();       /* 选 i 斗开 */
	sleep(1000);
	p53d1_1();      /* 启动炭黑慢加阀 */
   // if(jh77==1) {locat1(20,33);cor(0x1b,"粉料慢加");tlxk(i);}

	time(&t1);sum=0;
	while((ad+point)<weight)
	   {
	    while(th3==1) { p53d1_0();sleep(550); }
	      ma20();           /* 读A/D值 */
	      time(&t2);
	/*      while((difftime(t2,t1)>6)&&((ad)<1))     {
		   p53d1_0();
		   sleep(1000);
		ma20();   p54d5_1();
	      yterror(THD);ddc1();   p54d5_0();
	      } */
	      p53d1_1();
	   p55d_1();
	      con();
	      sleep(55);
	   if(ad>weight) ma20();
	   }
	 w2=ad;
	p55d_0();               /* i号斗关 */
	p53d1_0();              /* 关闭慢加阀 */
	tlxg(i);
	sleep(2500);
	ma20();w2=ad-w2;
  }
	w3=0;
	if((weight-ad)>error)   // 炭黑点动 
	{
	  jcc04=(int)(s1[m].sj*1000);
	   while((weight-ad)>error)   {
	    w3=ad;
	    while(th3==1)  sleep(3300);
	    p55d_1();           /* 选 i 斗开 */
	    sleep(1000);   
	    p53d1_1();          /* 启动慢加阀 */
	    tlxk(i);
	    sleep(jcc04);
	    p53d1_0();          /* 关闭慢加阀 */
	    p55d_0();           /* i号斗关 */
	    tlxg(i);
	    sleep(1000);
	    p55d_0();           /* i号斗关 */
	    while(th3==1)  sleep(3300);
	    sleep((int)(jcc[0][5]*1000));
	    ma20();
	    con();
	    w3=ad-w3;
	if(ad>weight) ma20(); 
	}
	}
	sleep(55);con();
	d_sz[d_t][0][m]=(ad-thor)/10.0;
	s1[m].sz=(ad-thor)/10.0+s1[m].sz;
/*      s1[m].tj=(1.5*w2+error)/100;
	s1[m].km=s1[m].tj+1.5*w1/100;
   #if COUNT==0
	locat1(0,1);
	printf("  w1=%5.2f w2=%5.2f w3=%5.2f",w1/100,w2/100,w3/100);
#endif */
	if(fabs(weight-ad)>error)    // 超公差
	{
	    p54d6_1();p19d2_0();p19d3_0();
	    yterror(THJ);
	    do
	      {
	       p19d3();i=nz;p19d2();
	       sleep(55);
	      }while((i|nz)==0);
	    p54d6_0();
	    if(nz==0)  {
			ma20();s1[m].sz=s1[m].sz-ad/10.0;thsl=-ad;m++;
	   if(m<(ms1+ms2)) goto done1; else {th=1;th1=1;m=0;goto go22;}
	   }
	}
	m++;
	if(m==ms1) pl(error);
	if(m<(ms1+ms2)) goto done1;
	/*time(&finish);
	if(difftime(finish,start)>((int)jcc[0][3])) yterror(THS);*/
	m=0;
	th=1;th1=1;
	if(ms2 != 0) { pl(error); th1=1; }
go22:;
	p19d0_0();p19d0();
	if(d_t == 0) d_t=1; else d_t=0;
	n++;
	if(ns>=n)
	  goto done;

	for(i=0;i<pfs1n;i++) {
	s1[i].sz=k22*s1[i].zl+s1[i].sz; }
/* #if COUNT==0
	
	lclcls(16,6,1,24);locat1(16,6);
	cor(0x8e,"碳黑秤完毕  ");
	locat1(17,2);cor(0x0a,"斗号  标准量  实耗量");
	for(i=0;i<pfs1n;i++) { locat1(17+i,2);n=s1[i].th;
	printf("%d 斗 %7.2f %7.2f",n,ns*s1[i].zl,s1[i].sz); }
	
#endif */
	do
	{
	  sleep(550);
	}while(ljmrun==1);
	goto go;
	/* p56d6_0(); */ s1run=1;
}


ma22()
{
float x;
	ad=io1(1)-tw_e1;
	/*s1l=ad/100.0;*/
}

ma20()  /* A/D值摸拟子程序 */
{
float x;int i;
	i=atoi(ljcha);
	if(ufha==1) i=-i;
	ad=i-tw_e1;
	/*s1l=ad/100.0;*/
}

pl(float error)
{
int nz;
 /*   por18=por18&lthr;sleep(880);
      if((por18&thr)!=thr) {
						#if COUNT==0
						    mov1();locat1(lclh,lcll);
						    printf("中间斗阀开故障");
						#else
						    mixerror(45);
						#endif
		  ddc1();  } */
	p16d3_0();p16d3();p19d4_0();
	if((nz!=0)||(th1==1)) {     // 等待卸料
	 }
	do {
	   p19d4();sleep(880);
	/* p16d3();
	   }while((th1==1)||(nz!=0)); */
	   }while((th1==1)&&(nz!=fiv));
	sleep(3300);
	p18d2_0();p18d2();
	if(nz!=thr) {yterror(45);ddc1();}
	do   {
	    p54d7_1();  /* 开称斗阀门,碳黑排下 */
	//  if(jh77==1){_setcolor(4);_ellipse(3,326,183,340,197);}
	    sleep(1500);
    	p18d1_0();p18d1();
    	if(nz!=two) {yterror(THD);/*p54d5_1();*/ ddc1();/*p54d5_0();*/}
	    sleep(5000);
	    p54d7_0();  /* 关碳黑称斗阀门 */
	//  if(jh77==1){_setcolor(15);_ellipse(3,326,183,340,197);}
	    sleep(1000);
	    ma20();
	   }while(ad>24);
	p54d7_1();      /* 开碳黑称斗阀门 */
	sleep(5000);
	p54d7_0();      /* 关称斗阀门 */
	sleep(1500);
	tw_e1=io1(1);
	ma20();
}

yterror(int x)
{
/* #if COUNT==0
	locat1(5,6);
	switch(x)
	{
	   case 45  :
		locat1(19,6);
		cor(0x8e,"中间斗阀故障    ");
		break;
	   case YLS :
		cor(0x8e,"油料秤超时      ");
		break;
	   case THS :
		cor(0x8e,"碳黑秤超时      ");
		break;
	   case YLG :
		locat1(11,6);
		cor(0x8e,"油料秤故障      ");
		break;
	   case THD :
		locat1(19,6);
		cor(0x8e,"碳黑秤故障      ");
		break;
	}
#else  */
	jherror(x);
// #endif
}
/* tlxk(int i)
{
 switch(i)
 {
  case 1:
		if(jh77==1){_setcolor(4);_ellipse(3,60,60,85,85);}
		break;
  case 2:
		if(jh77==1){_setcolor(4);_ellipse(3,151,60,176,85);}
		break;
  case 3:
		if(jh77==1){_setcolor(4);_ellipse(3,243,60,268,85);}
		break;
  case 4:
		if(jh77==1){_setcolor(4);_ellipse(3,410,60,435,85);}
		break;
  case 5:
		if(jh77==1){_setcolor(4);_ellipse(3,500,60,525,85);}
		break;
  case 6:
		if(jh77==1){_setcolor(4);_ellipse(3,590,60,615,85);}
		break;
 }
}
tlxg(int i)
{
 switch(i)
 {
  case 1:
		if(jh77==1){_setcolor(15);_ellipse(3,60,60,85,85);}
		break;
  case 2:
		if(jh77==1){_setcolor(15);_ellipse(3,151,60,176,85);}
		break;
  case 3:
		if(jh77==1){_setcolor(15);_ellipse(3,243,60,268,85);}
		break;
  case 4:
		if(jh77==1){_setcolor(15);_ellipse(3,410,60,435,85);}
		break;
  case 5:
		if(jh77==1){_setcolor(15);_ellipse(3,500,60,525,85);}
		break;
  case 6:
		if(jh77==1){_setcolor(15);_ellipse(3,590,60,615,85);}
		break;
 }
} */

ddc1()          // 炭黑秤等待继序运行
{
int nz;
	    p19d3_0();
	 do{ 
	    p19d3();sleep(880);
	   }while(nz!=fou);
}


ddc2()          // 油料秤等待继序运行
{
int nz;
	    p1ad4_0();
	 do{
	    p1ad4();sleep(880);
	   }while(nz!=fiv);
}


int h16_bcd(int i)
{
 int j,k,l;
 j=(int)(i/10.0);k=i-j*10;
 l=j;j=(int)(j/10.0);k=k+((l-10*j)<<4);
 l=j;j=(int)(j/10.0);k=k+((l-10*j)<<8);
 l=j;j=(int)(j/10.0);k=k+((l-10*j)<<12);
 return k;
}

int bcd_h16(int i)
{
 int j;
 j=1000*((i&0xf000)>>12);
 j=j+100*((i&0xf00)>>8);
 j=j+10*((i&0xf0)>>4);
 j=j+(i&0xf);
 printf(" j=%d \n",j);
 return j;
}

char *h16_b02(int i)
{
 int j=0;
 char p1[6];
 j=((i&0x8)<<9);
 j=j+((i&0x4)<<6);
 j=j+((i&0x2)<<3);
 j=j+(i&0x1);
 sprintf(p1,"%x",j);
// strcpy(d_3,p1);
 return p1;
}

char *ii_bin(unsigned int i,char *p)
{
 char p1[7];
 long j;
 j=i+10000;*p=0;
 sprintf(p1,"%ld",j);
 strcpy(p,p1+1);
 return p;
}

char *i_bin(unsigned int i,char *p)
{
 char p1[7];
 sprintf(p1,"%d",i);
 strcpy(p,p1);
 return p;
}

char *h16_10a(unsigned int i,char *p)
{
 int k,l;
 long j;
 char p1[10];
 j=i+0x10000;*p=0;
 sprintf(p1,"%lx",j);
 //strupr[p1];
 strcpy(p,p1+1);
 l=strlen(p);
 for(k=0;k<l;k++)
  if(*(p+k)>0x60) *(p+k)=*(p+k)-0x20;
 return p;
}

char *head_d(char *p,char *p1,unsigned int m,unsigned int i)
{
 int j,k,l;
 char p2[7];
 strcat(p,p1);
 j=m+10000;
 sprintf(p2,"%d",j);
 strcat(p,p2+1);
 *p2=0;
 j=i+0x100;
 sprintf(p2,"%x",j);
 for(k=0;k<3;k++)
  if(*(p2+k)>0x60) *(p2+k)=*(p2+k)-0x20;
 strcat(p,p2+1);
 return p;
}

/*void head_d(char *p,char *p1,unsigned int i)
{
 int j,k,l;
 char p2[10],p3[10];
 l=strlen(p1);
 for(j=0;j<5;j++) p2[j]='0';
 j=0;
 while(p1[j]>'9') {
  p2[j]=p1[j];
  j++;
 }
 for(k=l;k>j;k--)
  p2[5-k+j]=p1[j+l-k];
 p2[5]=0;
 strcat(p,p2);
 *p2=0;*p3=0;
 j=i+100;
 sprintf(p2,"%d",j);
 strcpy(p3,p2+1);
 strcat(p,p3);
// return p;
}*/

unsigned asc_10i(char *p)
{
 int i,l[4];
 for(i=0;i<4;i++)
  if(p[i]>'9') l[i]=(p[i]-0x37)&0xdf; else \
              l[i]=(p[i]-0x30)&0xdf;
 return(16*16*16*l[0]+16*16*l[1]+16*l[2]+l[3]);
}

unsigned asc_16i(char *p)
{
 return(10*10*10*(p[0]-0x30)+10*10*(p[1]-0x30)+\
        10*(p[2]-0x30)+(p[3]-0x30));
}

char *ad_d(char *p)
{
 int i,j;
 char p1;
 p1=0;
 j=strlen(p);
 for(i=1;i<j;i++) p1=p1+(*(p+i));
 if((p1&0xf0)>0x90) *(p+i)=0x37+((p1&0xf0)>>4);
 else *(p+i)=0x30+((p1&0xf0)>>4);
 if((p1&0xf)>0x9) *(p+i+1)=0x37+(p1&0xf);
 else *(p+i+1)=0x30+(p1&0xf);
 *(p+i+2)=0;
 return(p);
}

char *head_t(char *p,char *p1)
{
 *p=0x5;*(p+1)='0';*(p+2)='0';*(p+3)='F';*(p+4)='F';*(p+5)=0;
 strcat(p,p1);
 *(p+7)='0';*(p+8)=0;
 return p;
}

void bcom1(char *p)
{
 int j,k;
 bioscom(1,0x04,1);
 j=strlen(p);
 for(k=0;k<j;k++)
  bioscom(1,*(p+k),1);
}

void bcom2(char *p)
{
 int i,j,k;
 char p1[9];
 *p=0;
 do{
    j=bioscom(2,0,1);
   }while(j != 0x2);
 i=0;
 do{
    j=bioscom(2,0,1);
    if((j&0xff00)==0) d_2[i]=j;
    if((j&0xff00)!=0) printf(" j=%x \n",j);
    i++;
   }while(d_2[i-1] != 0x3);
 d_2[i-1]=0;
 j=strlen(d_2);
 k=(int)(j/4);
 for(i=1;i<k;i++)
  for(j=0;j<4;j++)
  {p_d[i][j]=d_2[i*4+j];
   p_d[i][j+1]=0;
  }
 for(k=1;k<i;k++)
  d_d[k]=asc_16i(p_d[k]);
 //return k;
}


char d_1[40],d_2[40],p_d[10][6];
char d_3[6],d_4[6],m_d[10][6];
int d_d[10];
void main()
{
 int i,j,k;
 char p1,p2,zoj[10],zo1[10];

 /*_bios_serialcom(_COM_INIT,1,0xfa);*/
 bioscom(0,0xea,1);

 head_t(d_1,"WW");
 i=0;
 head_d(d_1,"D",i,25);
 for(i=200;i<225;i++)
 strcat(d_1,ii_bin(i,zoj));
/* strcat(d_1,ii_bin(498,zoj));
 strcat(d_1,h16_10a(789,zoj));
 strcat(d_1,h16_10a(256,zoj));
 strcat(d_1,h16_10a(3689,zoj));*/
 ad_d(d_1);
 printf(" A=%s \n",d_1);
/* printf(" p=%x \n",h16_bcd(0xf));
 printf(" p=%d \n",bcd_h16(31));*/
 bcom1(d_1);
 head_t(d_1,"BW");
 i=800;j=1;
 head_d(d_1,"M",i,25);
 for(i=0;i<25;i++)
 {if(j==0) j=1; else j=0;
 strcat(d_1,i_bin(j,zoj));}
 ad_d(d_1);
 printf(" B=%s \n",d_1);
 bcom1(d_1);
 head_t(d_1,"WR");
 head_d(d_1,"D",0,5);
 ad_d(d_1);
 bcom1(d_1);
 bcom2(d_2);
 for(i=1;i<5;i++)
 printf(" k=%d \n",d_d[i]);
}