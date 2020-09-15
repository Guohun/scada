    char  temp[40],temp1[40];
    int i=0,j,k;
	float f1;
	char s1[40],s2[40];
	unsigned c1;

	HDC  hdc;
/*          x=sio_close(2);
          Sleep(200);
          x=sio_open(2);
		  if (x<0)
		  {
			   MessageBox(hwnd,"open  error","open 2",0);
               return  0;
		  }
            wsprintf(temp, "call sio_ioctl return value = %d", sio_ioctl(2,
				B2400, BIT_7 | STOP_2 | P_EVEN));

				for(i=0;i<8;i++)
                    temp1[i]=' ';
                 temp1[i]='\0'; */
              while (TRUE)
               {
                hdc=GetDC(hwnd);
  	/*		    sio_flush(2,0);	
                Sleep(100);

				do
				{
                       i=sio_getch(2);
    				   Sleep(10);
                }while(i!=2);

                sio_read(2,s1,9);
				c1=s1[0]&0x007;
				for(i=0;i<7;i++)
				   s2[i]=s1[i+3];

				k=atoi(s2);
				switch(c1)
				{
					case 0x01: f1=k*10.0;break;
					case 0x03: f1=k/10.0;break;
					case 0x04: f1=k/100.0;break;
					case 0x05: f1=k/1000.0;break;
					case 0x06: f1=k/10000.0;break;
				};
                   if((s1[1]&0x02)!=0) f1=-f1;
                   sprintf(temp1,"%8.4f",f1);
                   TextOut(hdc,10,200,temp1,8);
   				   sio_flush(2,0); */
                   Sleep(200);

               };
                ReleaseDC(hwnd,hdc);
             /*   x=sio_close(2);
                  if (x<0)
                  {
                       MessageBox(hwnd,"close  error","open",0);
                        return  0;
                 }*/

