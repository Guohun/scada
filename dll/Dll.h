/*
			98.4.22  �Ķ���¼ 
		  float			 temp2;						/* ��Ϊfloat �� ��ʾС�Ϲ���
//          int            temp2;                      /* ���� 
*/

  struct  pei_fan_table
  {
          int  index;         /*Ŀ¼���*/
          char   _number[8];       /*�䷽��� */
          char  name[15];     /* �䷽���� */
          char   day[12];
          char   time[12];
          char  ban;            /* ���*/
          char  che;           /*  ����*/
          char  flag;         /*   ��־     0 û����  1---���ڽ���  2---����� */
                                         /*����ʱ�Զ���Ϊ  0*/
          char  now_che;      /*��ǰ����  ��Ϊ0*/
          char  now_duan;      /*��ǰ����  ��Ϊ0*/

  };
    struct   lian_add_type
    {
          char           _number[8]        ;         /*  ��� */
          int            min_wen_du        ;         /* ���޽��� */
          int            max_wen_du        ;         /* ���޽��� */
          char           dou_liao_time     ;         /* Ͷ��ʱ�� */
          char           dou_liao_start    ;         /* ��������0   ���� */
          float          weight            ;         /* С������ */
          int            pai_liao_time;               /*�������� */
          int            pai_liao_ask;                /*����ʱ��*/
          int            temp1;                      /* ���� */
		  float			 temp2;						/* ��Ϊfloat ��*/
//          int            temp2;                      /* ���� */
          int            temp3;                     /* ���� */
     };

    struct   pei_fan_type
		   {
              char  _number[8];         /*�䷽���*/
              char  cai_number[8];      /* ���ϱ�� */
              char  name[18];            /*��������*/
              int   mathine;           /*�����    0    1 2 */
              char  wna[5];              /* ����  1
											,2
											,3,
											4
											*/
             char   dou[5];         /* ����     1  ���� 1,2
                                                    2    dou2
                                                    3    dou3
                                                */
            float    fast_do;              /* ������ǰ�� */
            float    drop_do;              /* �㶯��ǰ�� */
            float    control_time;         /* ����ʱ�� */
            int      add_time;              /*���ϴ��� */
            int      stop_time;             /*�ȶ�ʱ�� */

            float    weight;               /* �������� */
            float    gon_cha;              /* ������ */
            float    get_weight;           /*ʵ������  ӦΪ float*/
        };

        struct  hun_lian_type
        {
                  char  _number[8];         /*�䷽���*/
                  int    duan ;             /*�κ�  1 2 3 4*/
                  int    dou_you_time;    /* ��Ͷ��    Ϊ0 ��Ͷ*/
			      int    dou_you2_time;    /* ��Ͷ��2    Ϊ0 ��Ͷ ��*/         
                  int    dou_tan_time;    /* ��Ͷ̿   Ϊ 0 ��Ͷ*/
                  int    dou_jiao_time;   /* ��Ͷ��  Ϊ  0 ��Ͷ*/
                  int    dou_xiao_time;   /* ��ͶС��  Ϊ 0 ��Ͷ*/

                  int    st;              /* Ͷ������ 0 */
                  char       speed;            /* ת�� */
                  char       pre;              /* ѹ���� */
                  char       lian_time;        /* ����ʱ�� */
                  int        upper_pre;         /*�϶�˨ѹ�� */
                  int        tem;              /* �¶� */
                  float      neg;               /* ���� */
                  char       ctr;              /* ���ƹ�ϵ */
                  int	    stop_time;          /*�ȶ�ʱ�� */
        };

    struct Produce_type {
          char  sj1_sum;     /* ��ǰ��һ�ν��ϸ��� */
          char  sj1_max;     /* ��һ������ϸ��� */
          char  sj2_sum;     /* ��ǰ�ڶ��ν��ϸ��� */
          char  sj2_max;     /* �ڶ�������ϸ��� */
          char  th1_sum;     /* ��ǰ��һ�η��ϸ��� */
          char  th1_max;     /* ��һ�������ϸ��� */
          char  th2_sum;     /* ��ǰ�ڶ��η��ϸ��� */
          char  th2_max;     /* �ڶ��������ϸ��� */
          char  yz1_sum;     /* ��ǰ��һ����һ���� */
          char  yz1_max;     /* ��һ�������һ���� */
          char  yz2_sum;     /* ��ǰ�ڶ�����һ���� */
          char  yz2_max;     /* �ڶ��������һ���� */
          char  yt1_sum;     /* ��ǰ��һ���Ͷ����� */
          char  yt1_max;     /* ��һ������Ͷ����� */
          char  yt2_sum;     /* ��ǰ�ڶ����Ͷ����� */
          char  yt2_max;     /* �ڶ�������Ͷ����� */
          char  pei_sum;     /* ��ǰ���� */
          char  pei_max;     /* ������ */
          struct   lian_add_type  base;    /* �䷽�������� */
          struct   pei_fan_type  th1[10];  /* ̿��1�䷽ */
          struct   pei_fan_type  th2[10];  /* ̿��2�䷽ */
          struct   pei_fan_type  sj1[10];  /* ����1�䷽ */
          struct   pei_fan_type  sj2[10];  /* ����2�䷽ */
          struct   pei_fan_type  yz1[10];  /* ��һ1�䷽ */
          struct   pei_fan_type  yz2[10];  /* ��һ2�䷽ */
          struct   pei_fan_type  yt1[10];  /* �Ͷ�1�䷽ */
          struct   pei_fan_type  yt2[10];  /* �Ͷ�2�䷽ */
          struct   hun_lian_type huan[20];  /* �������䷽ */
    };

       struct run_table		/* vb-->dll --->VB */
       {
                short  total_che; /*�ܳ���*/
				short  now_che;/*�ڼ���*/
				short  all_duan;/* �ܶ� */
				short  ban;		/*��  */
				short  mathine;  /*���� */
				long  all_time;	  /*   ȡ�� ����*/
				long  all_che_time;	/*��һ��ʱ��  1 ��ʾ��ʼ  -1 ����*/
				long  duan_time; /*��ʱ�� 1 ��ʾ��ʼ  -1 ����*/
                char _number[8];  /*�䷽��� */
                char  name[18];   /*�䷽���� */
        };

            struct   make_base_type         /* �䷽�������� */
            {
                  char  sj1_sum;     /* ��ǰ��һ�ν��ϸ��� */
                  char  sj1_max;     /* ��һ������ϸ��� */
                  char  sj2_sum;     /* ��ǰ�ڶ��ν��ϸ��� */
                  char  sj2_max;     /* �ڶ�������ϸ��� */
                  char  th1_sum;     /* ��ǰ��һ�η��ϸ��� */
                  char  th1_max;     /* ��һ�������ϸ��� */
                  char  th2_sum;     /* ��ǰ�ڶ��η��ϸ��� */
                  char  th2_max;     /* �ڶ��������ϸ��� */
                  char  yz1_sum;     /* ��ǰ��һ����һ���� */
                  char  yz1_max;     /* ��һ�������һ���� */
                  char  yz2_sum;     /* ��ǰ�ڶ�����һ���� */
                  char  yz2_max;     /* �ڶ��������һ���� */
                  char  yt1_sum;     /* ��ǰ��һ���Ͷ����� */
                  char  yt1_max;     /* ��һ������Ͷ����� */
                  char  yt2_sum;     /* ��ǰ�ڶ����Ͷ����� */
                  char  yt2_max;     /* �ڶ�������Ͷ����� */

                  char  pei_max;     /* ������ */

                  struct   pei_fan_type  th1[10];  /* ̿��1�䷽ */
                  struct   pei_fan_type  th2[10];  /* ̿��2�䷽ */
                  struct   pei_fan_type  sj1[10];  /* ����1�䷽ */
                  struct   pei_fan_type  sj2[10];  /* ����2�䷽ */
                  struct   pei_fan_type  yz1[10];  /* ��һ1�䷽ */
                  struct   pei_fan_type  yz2[10];  /* ��һ2�䷽ */
                  struct   pei_fan_type  yt1[10];  /* �Ͷ�1�䷽ */
                  struct   pei_fan_type  yt2[10];  /* �Ͷ�2�䷽ */


                  float  get_weight;      /*С�� ʵ��  */
                  int  name;             /*  С������  */
                  int   get_pai_time;         /* ʵ������ʱ�� */

                  int   huan_time;             /*������ʱ��  */
                  int   pai_tempro;            /*����ʱ�¶�  */
                  float   all_power;            /*����������  */
            };

            struct   make_lian_liao              /*ֻ�����  ÿ�� д��make_lian_liao table*/
             {
                  int    mathine;
                  long    only_key;       /*�ؼ��� Ψһ  */
                  int    duan_hao;      /* �κ� */
                  int    now_che;      /*��ǰ����  */

                  int    get_hun_time;        /* ʵ�ʻ���ʱ��      �Ʊ� �� */
                  int    set_hun_time;         /* �趨����ʱ��     �Ʊ� �� */

                  int   set_you_time;       /* ��Ͷ��    Ϊ0 ��Ͷ   �Ʊ� �� */
			      int   set_you2_time;    /* ��Ͷ��2    Ϊ0 ��Ͷ ��*/         

                  int   set_tan_time;       /* ��Ͷ̿   Ϊ 0 ��Ͷ   �Ʊ� ��  */
                  int   set_jiao_time;      /* ��Ͷ��  Ϊ  0 ��Ͷ   �Ʊ� ��  */
                  int   set_xiao_time;      /* ��ͶС��  Ϊ 0 ��Ͷ  �Ʊ� ��  */

                  int   get_you_time;       /* ��Ͷ��    Ϊ0 ��Ͷ  �Ʊ� ��   */
                  int   get_tan_time;       /* ��Ͷ̿   Ϊ 0 ��Ͷ  �Ʊ� ��	*/
                  int   get_jiao_time;      /* ��Ͷ��  Ϊ  0 ��Ͷ  �Ʊ� ��  */
                  int   get_xiao_time;      /* ��ͶС��  Ϊ 0 ��Ͷ �Ʊ� ��  */
			      int   get_you2_time;    /* ��Ͷ��2    Ϊ0 ��Ͷ ��*/         

                  int   next_tempro;            /*ת�¶�����---- �Ʊ� ��ͼ�� */
                  float   next_power;             /*ת�¶ι���---- �Ʊ� ��ͼ��*/

                  char  duan_begin_time[14];         /*�ο�ʼʱ��     �Ʊ� ��ͼ�� */
                  char  duan_end_time[14];          /*�ν���ʱ��      �Ʊ� ��ͼ��*/

                  char  power[3000];  /*����ֵ  ��Ÿ�ʽ: 30 40 60 memo */
                  char  tempro[3000];  /*�¶�ֵ*/

                  int  set_termpro;
                  float  set_power;
                  int  set_speed;

                  char  Chian_Str[20];      /*��һ�ּ�¼��ʽ*/
                  char  Eng_Str[20];
                  int  Set_the_time;
                  int  Get_the_time;
              };



        struct power_tempro				//�¶ȼ�¼
        {
                int duan_hao;
                int flag;
                int total;
                int ptr;
                float power[600];		//����
                int tempro[600];
        };

/*
        elec_flag �ĺ��壺
                ����----����
                2----����                
                3----����ȷ�Ϻ� �Զ�Ϊ0                
*/

                struct  elect_send_type	//���źż�¼
                {
					short    flag;
                    short    length;         //  ʵ�ʳ���
                    short    note_name[501]; //     �ڵ���
                    short    data[501];     // ����
                  };

		struct  dou_type	//����¼
		{
			short  length;
			short  flag[40];
			short  mathine[40];
			short  name[40];			
		};
		struct  ban_type  //�Ƽ�¼
		{
			short  length;
			float   max[20];
			float   weight[20];
			short   mathine[20];			//����
			short    name[20];
			short    flag[20];
		};
		struct 	G_error_num
		{			
			short  length;
            long   data[101];
            char   prompt[101][100];
		};

struct  screen_buff	/*��������  */
{
	int  Clear_flag;
	int  change_flag;
	int row;
	int  col;
	char str[100];
};

/*ͨѶ�� */
	struct send_table
	{		
	      int  code;      /*���� */
           char name[20];    /*���ڶ�Ӧ�߼���*/
              float  min;       /*��Сֵ*/
              float  max;       /*���ֵ*/
              int    AD_max;     /* ���ADֵ*/
              int    AD_min;     /* ��СADֵ*/
              char   port;   /*�ڵ�ַ*/
              float  time;      /*����ʱ��*/
              float   true_data; /*ʵ�ʻ�������*/
              int   mathine;    /*������*/
              int     true_AD;   /*ʵ��AD*/
	      char      state;
              char     run_status   ;       //=1����  ��0 ��
	};
		
//�������ã�
struct	elec_in_table
	{
	 int	name;//��������
	 char   str[6];//����	
	 int	run;	//�������ӵ�
	 int    run_set; 		//�����ӵ��߼��趨ֵ
	 int	check;			//��
	 int    check_set;		//�߼��趨ֵ
	 int    run_value;		//ʵ��ֵ
	 int    check_value;
	 int	check_time;		//ʱ��
	 int	use_flag;		//ͣ���־
	};

/*******************
	����ϵͳ
********************/
       struct Fs_ml_table		/*������������ */
       {
				int    Fm_ch;	/*���*/
				int    Fm_jzh;	/*�����*/
				int    Fm_dh;	/* ����*/
				int    Fm_pzh;	/* Ʒ�ֺ�*/
		                char   Fm_clbh[8];  /*���ϱ�� */
				char   Fm_clname[18];/*�������� */
				int    Fm_sgs;		/*���͹���*/	
				int    Fm_lsmgs;	/*����������*/	
				int    Fm_bl;		/*����*/
        };

       struct Fs_sb_table		/* �����豸������ */
       {
				int    Fsb_jzh;	/*���*/
				int    Fsb_dh;	/* ����*/
				int    Fsb_pzh;	/* Ʒ�ֺ�*/
		                float   Fsb_dwight;  /*�������� */
				float   Fsb_gzyl;    /*����ѹ�� */
				float    Fsb_zysd;    /*��ѹ�趨*/	
				float    Fsb_dysd;    /*��ѹ�趨*/	
				float    Fsb_dsgy;    /*������ѹ*/
				int    Fsb_mtime;     /*�������ʱ��  */		
				int	Fsb_qtime;    /*��ɨʱ��  */
				int	Fsb_ty;	      /*ͣ��  */
				float	Fsb_blxs;     /*����ϵ��  */
				int	Fsb_byl;      /*����   */
        };

	struct  ban_table
	{
		char _number[8];
		char name[18];
		short ban;
		short _che;
		short mathine;
		short  total_che;
	};



#define CCONV _stdcall
#define NOMANGLE
#define  CAN_READ  1
#define  CAN_RUN	1
#define  SIMULATE   2
#define  STOP		0
#define  STOP_DOING  0
#define  WRITE_DOING  -1
#define  MAX_DOU   20
#define  MAX_BAN   8


#ifdef __cplusplus
extern "C" {            /* Assume C declarations for C++ */
#endif  /* __cplusplus */

/*�˴�ģ����  */
#ifdef _WIN32
	#ifdef _ZGH_USER_
		#define	WINMMAPI	DECLSPEC_IMPORT
	#else
		#define	WINMMAPI	
	#endif
#endif

WINMMAPI int WINAPI read_fen_ban(struct  ban_table *x);
WINMMAPI int WINAPI read_grap(char *x);
WINMMAPI int WINAPI read_Fsml_table(struct Fs_ml_table *x,int *flag,int *number);
WINMMAPI int WINAPI write_Fsml_table(struct Fs_ml_table *x,int flag,int number);
WINMMAPI int WINAPI read_Fssb_table(int mathine,struct Fs_sb_table   *x);
	
WINMMAPI int WINAPI write_chans(int,int,int,int);
     //	����š����š�Ʒ�ֺš����͹���
WINMMAPI int WINAPI write_runfch(int,float,float,float,int,char *);
     //	��ǰ������1#��ѹ����2#��ѹ�����ܵ�ѹ����ʱ�䣬״̬



/*�������У�ģ�⴦����:�����ݲ��ֺ���  */
/*
        run_flag �ĺ��壺
               -1---��������
                0----ֹͣ
                1----��ʼ����
                2----ģ��
                3----�����޸�
                4----�����޸�ȷ�� ��run_flag==1

*/
WINMMAPI int WINAPI  read_produce(struct Produce_type  *temp);
WINMMAPI int WINAPI  read_run(struct run_table* x,int *run_flag);
WINMMAPI int WINAPI  write_run(struct run_table* x,int run_flag);
WINMMAPI int WINAPI  init_elec_input(struct elect_send_type *x);	//д��
WINMMAPI int WINAPI  init_elec_output(struct elect_send_type *x);	//д����
WINMMAPI int WINAPI Set_Turn(int elec_name,int data);
WINMMAPI int WINAPI Set_Light(int elec_name,int data);

/*�������У�ģ�⴦����:д���ݲ��ֺ���  */
WINMMAPI int WINAPI  write_elec_input(struct elect_send_type *x);
WINMMAPI int WINAPI  write_elec_output(struct elect_send_type *x);
WINMMAPI int WINAPI  write_Tan_dou(int mathine,int  dou_hao,int  flag);
WINMMAPI int WINAPI  write_ri_chu_dou(int mathine,int  dou_hao,int  flag);
WINMMAPI int WINAPI  write_You_dou(int mathine,int  dou_hao,int  flag);
//д��ǰ������
/*
        mathine  ����
        name     �Ӻ�
        max       û����
        weight   ����
        flag    �ͣ�̼
                =1  û����
                =2  ���
                =3  ����
                =4  �㶯
                =5  �ȴ�
                ��      =1 �齺  =2  Ƭ��

*/
WINMMAPI int WINAPI  write_ban(int mathine,int name,float far *max,
                float  far *weight,int flag);


/*���´�������  */
WINMMAPI int WINAPI write_error(int  error_num,char *prompt);
WINMMAPI int WINAPI read_error(short  error_num);
WINMMAPI int WINAPI kill_error(int error_num);
WINMMAPI int WINAPI Get_Elec_Check_Table(struct	elec_in_table x[],int *n);

/*���¶�������  */
WINMMAPI int WINAPI write_multi_screen(int row,int col,char *str);
WINMMAPI int WINAPI Clear_multi_screen(void);
WINMMAPI int WINAPI Init_Font(int size,char * name);

/*�����豸��  */
WINMMAPI int WINAPI read_send_table(int code,int mathine,struct send_table *x);
WINMMAPI int WINAPI  Z_DLL_Close_All();
WINMMAPI int WINAPI  Sio_Open_All( );
WINMMAPI int WINAPI  Read_VbKey();             


/*��������ͳ����  */
//����һ�� ����
WINMMAPI int WINAPI write_che(struct make_base_type temp);

//����һ�� ����
WINMMAPI int WINAPI write_make(struct make_lian_liao make_input);
WINMMAPI int WINAPI  write_power(int,float power,int tempro,int);


/*����ģ����  */
WINMMAPI int WINAPI  P_DO(void);
WINMMAPI int WINAPI  V_DO(void);
WINMMAPI int WINAPI  Init_P(void);
WINMMAPI int WINAPI  Init_V(void);
WINMMAPI int WINAPI  write_select_ban(int mathine,char *temp);

#ifdef __cplusplus
}                       /* End of extern "C" { */
#endif  /* __cplusplus */

