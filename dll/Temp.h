#define ODBC_ERROR       2001


extern int  Get_Date(char *buffer );
extern int  Get_Time(char *buffer );
static struct make_index		//一配方依次
{
	long  Sort_Idx;
	int   mathine;
	char  m_number[6];
	char  name[20];
	int   ban;
	int   che;
	int   flag;
	int   chang_flag;
	char  begin_date[20];
	char  begin_time[20];

	char    end_date[20];
	char    end_time[20];

};
/*
         struct  make_table         //每车一次 dll->库 
         {
                  long     only_key;       /*关键字 唯一  生产完一个记录 
                  char     m_number[8];       /*配方编号 
                  char     name[18];     /* 配方名称 
                  int      ban;     //        班号
                  int      che;   //          车数
                  int      mathine;         /*机组号
                  int      now_che;        /*当前车数
                  char     now_duan;      /*当前段数 
                  int      flag;    //        标志     0 没进行  1---正在进行  2---已完成
                  char     m_date[14];
                  char     m_time[14];
        };

*/

static struct make_table		//一车记录
{
	long  Sort_Idx;
	long  only_key;
	int   mathine;
	int   now_che;
	int   chang_flag;
	char  begin_date[20];  
	char  begin_time[20];  
	char  end_date[20];
	char  end_time[20];
};

struct  make_base         /* 配方基本数据 */
{
                  long   only_key;       /*关键字 唯一  */
                  int   now_che;      /*当前车数  */

                  float  set_weight;                /* 小料重量 */
                  float  get_weight;               /*小料 实重  */
                  int  name     ;                 /*  小料名称  */
                  float  gon_cha  ;                  /*小料公差    */
                  int  set_pai_time;         /* 设定排料时间       制表 绘图用*/
                  int  get_pai_time;         /* 实际排料时间       制表 绘图用*/
                  int  pai_tempro ;            /*排料时温度             制表 绘图用  */
                  int  huan_time   ;          /*混炼总时间             制表 绘图用   */
                  float  all_power;               /*消耗总能量   ---->统计得  */
};

struct  Fen_Disp_Dou
{
	int mathine;
	int dou;
	int kind;
	int  send_data;

	int  now_data;
	unsigned float  one;
	unsigned float  two;
	unsigned float    power;
	int	   my_time;
	char   statu[40];
};


	#pragma     data_seg("sdata")
      
     struct power_tempro   work_power;
     struct elect_send_type   elect_input;
       struct elect_send_type   elect_output;
       struct  G_error_num G_error;
       struct dou_type GT_dou;
       struct dou_type GY_dou;
       struct dou_type GR_dou;
       struct ban_type G_ban;
       CRITICAL_SECTION m_csMyLock;
       HANDLE  my_envent;
       HANDLE  main_thread;
       struct run_table    send_run_table;
       struct Produce_type  send_pei_fan;
	   struct make_index  Static_make_index;
       struct  make_table  Static_make_table;
       int  control_can_read;
       int  can_run_flag;
       int  g_mathine;
        char  duan_begin_time[20];
        char  Grap_File_Name[20];
        int   FONT_SIZE;
        HFONT  hNewf;
		int   comm_mdb_open_flag; /* 0---没打开* 1--打开*/
		int  make_mdb_open_flag;
		int  ljxt_mdb_open_flag;
		HENV make_henv;
	    HDBC make_hdbc; 
		HENV comm_henv;
		HDBC comm_hdbc; 		    
	    HENV ljxt_henv;
	    HDBC ljxt_hdbc; 
		int  port_array[8];
		struct  Fen_Disp_Dou  Chans_Data;
		int  Pai_Fan_Run_Flag;
		char grap_name[20];
	#pragma data_seg()
