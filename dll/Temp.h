#define ODBC_ERROR       2001


extern int  Get_Date(char *buffer );
extern int  Get_Time(char *buffer );
static struct make_index		//һ�䷽����
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
         struct  make_table         //ÿ��һ�� dll->�� 
         {
                  long     only_key;       /*�ؼ��� Ψһ  ������һ����¼ 
                  char     m_number[8];       /*�䷽��� 
                  char     name[18];     /* �䷽���� 
                  int      ban;     //        ���
                  int      che;   //          ����
                  int      mathine;         /*�����
                  int      now_che;        /*��ǰ����
                  char     now_duan;      /*��ǰ���� 
                  int      flag;    //        ��־     0 û����  1---���ڽ���  2---�����
                  char     m_date[14];
                  char     m_time[14];
        };

*/

static struct make_table		//һ����¼
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

struct  make_base         /* �䷽�������� */
{
                  long   only_key;       /*�ؼ��� Ψһ  */
                  int   now_che;      /*��ǰ����  */

                  float  set_weight;                /* С������ */
                  float  get_weight;               /*С�� ʵ��  */
                  int  name     ;                 /*  С������  */
                  float  gon_cha  ;                  /*С�Ϲ���    */
                  int  set_pai_time;         /* �趨����ʱ��       �Ʊ� ��ͼ��*/
                  int  get_pai_time;         /* ʵ������ʱ��       �Ʊ� ��ͼ��*/
                  int  pai_tempro ;            /*����ʱ�¶�             �Ʊ� ��ͼ��  */
                  int  huan_time   ;          /*������ʱ��             �Ʊ� ��ͼ��   */
                  float  all_power;               /*����������   ---->ͳ�Ƶ�  */
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
		int   comm_mdb_open_flag; /* 0---û��* 1--��*/
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
