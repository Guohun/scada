/*
			98.4.22  改动记录 
		  float			 temp2;						/* 改为float 型 表示小料公差
//          int            temp2;                      /* 备用 
*/

  struct  pei_fan_table
  {
          int  index;         /*目录序号*/
          char   _number[8];       /*配方编号 */
          char  name[15];     /* 配方名称 */
          char   day[12];
          char   time[12];
          char  ban;            /* 班号*/
          char  che;           /*  车数*/
          char  flag;         /*   标志     0 没进行  1---正在进行  2---已完成 */
                                         /*输入时自动设为  0*/
          char  now_che;      /*当前车数  设为0*/
          char  now_duan;      /*当前车数  设为0*/

  };
    struct   lian_add_type
    {
          char           _number[8]        ;         /*  编号 */
          int            min_wen_du        ;         /* 下限胶温 */
          int            max_wen_du        ;         /* 上限胶温 */
          char           dou_liao_time     ;         /* 投料时间 */
          char           dou_liao_start    ;         /* 排料上升0   备用 */
          float          weight            ;         /* 小料重量 */
          int            pai_liao_time;               /*排料条件 */
          int            pai_liao_ask;                /*排料时间*/
          int            temp1;                      /* 备用 */
		  float			 temp2;						/* 改为float 型*/
//          int            temp2;                      /* 备用 */
          int            temp3;                     /* 备用 */
     };

    struct   pei_fan_type
		   {
              char  _number[8];         /*配方编号*/
              char  cai_number[8];      /* 材料编号 */
              char  name[18];            /*材料名称*/
              int   mathine;           /*机组号    0    1 2 */
              char  wna[5];              /* 秤名  1
											,2
											,3,
											4
											*/
             char   dou[5];         /* 斗号     1  胶料 1,2
                                                    2    dou2
                                                    3    dou3
                                                */
            float    fast_do;              /* 快慢提前量 */
            float    drop_do;              /* 点动提前量 */
            float    control_time;         /* 控制时间 */
            int      add_time;              /*加料次数 */
            int      stop_time;             /*稳定时间 */

            float    weight;               /* 所需重量 */
            float    gon_cha;              /* 允许公差 */
            float    get_weight;           /*实际重量  应为 float*/
        };

        struct  hun_lian_type
        {
                  char  _number[8];         /*配方编号*/
                  int    duan ;             /*段号  1 2 3 4*/
                  int    dou_you_time;    /* 所投油    为0 不投*/
			      int    dou_you2_time;    /* 所投油2    为0 不投 补*/         
                  int    dou_tan_time;    /* 所投炭   为 0 不投*/
                  int    dou_jiao_time;   /* 所投胶  为  0 不投*/
                  int    dou_xiao_time;   /* 所投小料  为 0 不投*/

                  int    st;              /* 投料上升 0 */
                  char       speed;            /* 转速 */
                  char       pre;              /* 压力级 */
                  char       lian_time;        /* 混炼时间 */
                  int        upper_pre;         /*上顶栓压力 */
                  int        tem;              /* 温度 */
                  float      neg;               /* 能量 */
                  char       ctr;              /* 控制关系 */
                  int	    stop_time;          /*稳定时间 */
        };

    struct Produce_type {
          char  sj1_sum;     /* 当前第一次胶料个数 */
          char  sj1_max;     /* 第一次最大胶料个数 */
          char  sj2_sum;     /* 当前第二次胶料个数 */
          char  sj2_max;     /* 第二次最大胶料个数 */
          char  th1_sum;     /* 当前第一次粉料个数 */
          char  th1_max;     /* 第一次最大粉料个数 */
          char  th2_sum;     /* 当前第二次粉料个数 */
          char  th2_max;     /* 第二次最大粉料个数 */
          char  yz1_sum;     /* 当前第一次油一个数 */
          char  yz1_max;     /* 第一次最大油一个数 */
          char  yz2_sum;     /* 当前第二次油一个数 */
          char  yz2_max;     /* 第二次最大油一个数 */
          char  yt1_sum;     /* 当前第一次油二个数 */
          char  yt1_max;     /* 第一次最大油二个数 */
          char  yt2_sum;     /* 当前第二次油二个数 */
          char  yt2_max;     /* 第二次最大油二个数 */
          char  pei_sum;     /* 当前段数 */
          char  pei_max;     /* 最大段数 */
          struct   lian_add_type  base;    /* 配方基本数据 */
          struct   pei_fan_type  th1[10];  /* 炭黑1配方 */
          struct   pei_fan_type  th2[10];  /* 炭黑2配方 */
          struct   pei_fan_type  sj1[10];  /* 胶料1配方 */
          struct   pei_fan_type  sj2[10];  /* 胶料2配方 */
          struct   pei_fan_type  yz1[10];  /* 油一1配方 */
          struct   pei_fan_type  yz2[10];  /* 油一2配方 */
          struct   pei_fan_type  yt1[10];  /* 油二1配方 */
          struct   pei_fan_type  yt2[10];  /* 油二2配方 */
          struct   hun_lian_type huan[20];  /* 炼胶段配方 */
    };

       struct run_table		/* vb-->dll --->VB */
       {
                short  total_che; /*总车数*/
				short  now_che;/*第几车*/
				short  all_duan;/* 总段 */
				short  ban;		/*班  */
				short  mathine;  /*机号 */
				long  all_time;	  /*   取消 不用*/
				long  all_che_time;	/*总一车时间  1 表示开始  -1 结束*/
				long  duan_time; /*段时间 1 表示开始  -1 结束*/
                char _number[8];  /*配方编号 */
                char  name[18];   /*配方名称 */
        };

            struct   make_base_type         /* 配方基本数据 */
            {
                  char  sj1_sum;     /* 当前第一次胶料个数 */
                  char  sj1_max;     /* 第一次最大胶料个数 */
                  char  sj2_sum;     /* 当前第二次胶料个数 */
                  char  sj2_max;     /* 第二次最大胶料个数 */
                  char  th1_sum;     /* 当前第一次粉料个数 */
                  char  th1_max;     /* 第一次最大粉料个数 */
                  char  th2_sum;     /* 当前第二次粉料个数 */
                  char  th2_max;     /* 第二次最大粉料个数 */
                  char  yz1_sum;     /* 当前第一次油一个数 */
                  char  yz1_max;     /* 第一次最大油一个数 */
                  char  yz2_sum;     /* 当前第二次油一个数 */
                  char  yz2_max;     /* 第二次最大油一个数 */
                  char  yt1_sum;     /* 当前第一次油二个数 */
                  char  yt1_max;     /* 第一次最大油二个数 */
                  char  yt2_sum;     /* 当前第二次油二个数 */
                  char  yt2_max;     /* 第二次最大油二个数 */

                  char  pei_max;     /* 最大段数 */

                  struct   pei_fan_type  th1[10];  /* 炭黑1配方 */
                  struct   pei_fan_type  th2[10];  /* 炭黑2配方 */
                  struct   pei_fan_type  sj1[10];  /* 胶料1配方 */
                  struct   pei_fan_type  sj2[10];  /* 胶料2配方 */
                  struct   pei_fan_type  yz1[10];  /* 油一1配方 */
                  struct   pei_fan_type  yz2[10];  /* 油一2配方 */
                  struct   pei_fan_type  yt1[10];  /* 油二1配方 */
                  struct   pei_fan_type  yt2[10];  /* 油二2配方 */


                  float  get_weight;      /*小料 实重  */
                  int  name;             /*  小料名称  */
                  int   get_pai_time;         /* 实际排料时间 */

                  int   huan_time;             /*混炼总时间  */
                  int   pai_tempro;            /*排料时温度  */
                  float   all_power;            /*消耗总能量  */
            };

            struct   make_lian_liao              /*只须计算  每段 写入make_lian_liao table*/
             {
                  int    mathine;
                  long    only_key;       /*关键字 唯一  */
                  int    duan_hao;      /* 段号 */
                  int    now_che;      /*当前车数  */

                  int    get_hun_time;        /* 实际混料时间      制表 用 */
                  int    set_hun_time;         /* 设定混料时间     制表 用 */

                  int   set_you_time;       /* 所投油    为0 不投   制表 用 */
			      int   set_you2_time;    /* 所投油2    为0 不投 补*/         

                  int   set_tan_time;       /* 所投炭   为 0 不投   制表 用  */
                  int   set_jiao_time;      /* 所投胶  为  0 不投   制表 用  */
                  int   set_xiao_time;      /* 所投小料  为 0 不投  制表 用  */

                  int   get_you_time;       /* 所投油    为0 不投  制表 用   */
                  int   get_tan_time;       /* 所投炭   为 0 不投  制表 用	*/
                  int   get_jiao_time;      /* 所投胶  为  0 不投  制表 用  */
                  int   get_xiao_time;      /* 所投小料  为 0 不投 制表 用  */
			      int   get_you2_time;    /* 所投油2    为0 不投 补*/         

                  int   next_tempro;            /*转下段能量---- 制表 绘图用 */
                  float   next_power;             /*转下段功率---- 制表 绘图用*/

                  char  duan_begin_time[14];         /*段开始时间     制表 绘图用 */
                  char  duan_end_time[14];          /*段结束时间      制表 绘图用*/

                  char  power[3000];  /*能量值  存放格式: 30 40 60 memo */
                  char  tempro[3000];  /*温度值*/

                  int  set_termpro;
                  float  set_power;
                  int  set_speed;

                  char  Chian_Str[20];      /*另一种记录格式*/
                  char  Eng_Str[20];
                  int  Set_the_time;
                  int  Get_the_time;
              };



        struct power_tempro				//温度记录
        {
                int duan_hao;
                int flag;
                int total;
                int ptr;
                float power[600];		//能量
                int tempro[600];
        };

/*
        elec_flag 的含义：
                其余----正常
                2----调试                
                3----调试确认后， 自动为0                
*/

                struct  elect_send_type	//电信号记录
                {
					short    flag;
                    short    length;         //  实际长度
                    short    note_name[501]; //     节点名
                    short    data[501];     // 数据
                  };

		struct  dou_type	//斗记录
		{
			short  length;
			short  flag[40];
			short  mathine[40];
			short  name[40];			
		};
		struct  ban_type  //称记录
		{
			short  length;
			float   max[20];
			float   weight[20];
			short   mathine[20];			//决定
			short    name[20];
			short    flag[20];
		};
		struct 	G_error_num
		{			
			short  length;
            long   data[101];
            char   prompt[101][100];
		};

struct  screen_buff	/*多屏卡用  */
{
	int  Clear_flag;
	int  change_flag;
	int row;
	int  col;
	char str[100];
};

/*通讯用 */
	struct send_table
	{		
	      int  code;      /*代号 */
           char name[20];    /*串口对应逻辑名*/
              float  min;       /*最小值*/
              float  max;       /*最大值*/
              int    AD_max;     /* 最大AD值*/
              int    AD_min;     /* 最小AD值*/
              char   port;   /*口地址*/
              float  time;      /*运行时间*/
              float   true_data; /*实际回送数据*/
              int   mathine;    /*机器号*/
              int     true_AD;   /*实际AD*/
	      char      state;
              char     run_status   ;       //=1可用  ＝0 不
	};
		
//错误处理用：
struct	elec_in_table
	{
	 int	name;//错误名称
	 char   str[6];//保留	
	 int	run;	//阀启动接点
	 int    run_set; 		//启动接点逻辑设定值
	 int	check;			//灯
	 int    check_set;		//逻辑设定值
	 int    run_value;		//实际值
	 int    check_value;
	 int	check_time;		//时间
	 int	use_flag;		//停检标志
	};

/*******************
	风送系统
********************/
       struct Fs_ml_table		/*风送生产命令 */
       {
				int    Fm_ch;	/*序号*/
				int    Fm_jzh;	/*机组号*/
				int    Fm_dh;	/* 斗号*/
				int    Fm_pzh;	/* 品种号*/
		                char   Fm_clbh[8];  /*材料编号 */
				char   Fm_clname[18];/*材料名称 */
				int    Fm_sgs;		/*输送罐数*/	
				int    Fm_lsmgs;	/*连送最大罐数*/	
				int    Fm_bl;		/*备用*/
        };

       struct Fs_sb_table		/* 风送设备参数表 */
       {
				int    Fsb_jzh;	/*序号*/
				int    Fsb_dh;	/* 斗号*/
				int    Fsb_pzh;	/* 品种号*/
		                float   Fsb_dwight;  /*单罐重量 */
				float   Fsb_gzyl;    /*工作压力 */
				float    Fsb_zysd;    /*中压设定*/	
				float    Fsb_dysd;    /*低压设定*/	
				float    Fsb_dsgy;    /*堵塞高压*/
				int    Fsb_mtime;     /*最大输送时间  */		
				int	Fsb_qtime;    /*清扫时间  */
				int	Fsb_ty;	      /*停用  */
				float	Fsb_blxs;     /*比例系数  */
				int	Fsb_byl;      /*备用   */
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

/*此处模拟用  */
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
     //	机组号、斗号、品种号、输送罐数
WINMMAPI int WINAPI write_runfch(int,float,float,float,int,char *);
     //	当前罐数，1#罐压力，2#罐压力，管道压力，时间，状态



/*以下运行，模拟处理用:读数据部分函数  */
/*
        run_flag 的含义：
               -1---正在运行
                0----停止
                1----开始运行
                2----模拟
                3----在线修改
                4----在线修改确认 后run_flag==1

*/
WINMMAPI int WINAPI  read_produce(struct Produce_type  *temp);
WINMMAPI int WINAPI  read_run(struct run_table* x,int *run_flag);
WINMMAPI int WINAPI  write_run(struct run_table* x,int run_flag);
WINMMAPI int WINAPI  init_elec_input(struct elect_send_type *x);	//写灯
WINMMAPI int WINAPI  init_elec_output(struct elect_send_type *x);	//写开关
WINMMAPI int WINAPI Set_Turn(int elec_name,int data);
WINMMAPI int WINAPI Set_Light(int elec_name,int data);

/*以下运行，模拟处理用:写数据部分函数  */
WINMMAPI int WINAPI  write_elec_input(struct elect_send_type *x);
WINMMAPI int WINAPI  write_elec_output(struct elect_send_type *x);
WINMMAPI int WINAPI  write_Tan_dou(int mathine,int  dou_hao,int  flag);
WINMMAPI int WINAPI  write_ri_chu_dou(int mathine,int  dou_hao,int  flag);
WINMMAPI int WINAPI  write_You_dou(int mathine,int  dou_hao,int  flag);
//写当前秤重量
/*
        mathine  机号
        name     秤号
        max       没有用
        weight   重量
        flag    油，碳
                =1  没有用
                =2  快加
                =3  慢加
                =4  点动
                =5  等待
                胶      =1 块胶  =2  片胶

*/
WINMMAPI int WINAPI  write_ban(int mathine,int name,float far *max,
                float  far *weight,int flag);


/*以下错误处理用  */
WINMMAPI int WINAPI write_error(int  error_num,char *prompt);
WINMMAPI int WINAPI read_error(short  error_num);
WINMMAPI int WINAPI kill_error(int error_num);
WINMMAPI int WINAPI Get_Elec_Check_Table(struct	elec_in_table x[],int *n);

/*以下多屏卡用  */
WINMMAPI int WINAPI write_multi_screen(int row,int col,char *str);
WINMMAPI int WINAPI Clear_multi_screen(void);
WINMMAPI int WINAPI Init_Font(int size,char * name);

/*以下设备用  */
WINMMAPI int WINAPI read_send_table(int code,int mathine,struct send_table *x);
WINMMAPI int WINAPI  Z_DLL_Close_All();
WINMMAPI int WINAPI  Sio_Open_All( );
WINMMAPI int WINAPI  Read_VbKey();             


/*以下生产统计用  */
//生产一车 调用
WINMMAPI int WINAPI write_che(struct make_base_type temp);

//生产一段 调用
WINMMAPI int WINAPI write_make(struct make_lian_liao make_input);
WINMMAPI int WINAPI  write_power(int,float power,int tempro,int);


/*以下模拟用  */
WINMMAPI int WINAPI  P_DO(void);
WINMMAPI int WINAPI  V_DO(void);
WINMMAPI int WINAPI  Init_P(void);
WINMMAPI int WINAPI  Init_V(void);
WINMMAPI int WINAPI  write_select_ban(int mathine,char *temp);

#ifdef __cplusplus
}                       /* End of extern "C" { */
#endif  /* __cplusplus */

