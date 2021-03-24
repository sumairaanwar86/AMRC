#include "mbed.h"
#include "C12832.h"
#define SetKp 0x01                     //register where Kp will be stored
#define SetKi 0x02                     //register where Ki will be stored
#define SetKd 0x03                     //register where Kd will be stored
#define period 0.01 // ticker will read control signal from C&M every 0.01 sec

Ticker readdata;            // ticker that will read data from C&M device

DigitalOut ledon(LED1);    //LED that will show control and monitoring device is on
DigitalOut ledrun(LED2);   // LED that will show control and monitoring device is running

DigitalOut mode_s1(p21);       // mode selection signal to C&M device 1
DigitalOut mode_s2(p22);       // mode selection signal to C&M device 2
DigitalOut mode_s3(p23);       // mode selection signal to C&M device 3

AnalogIn PID_read1(p15);       // control signal  from C&M device 1
AnalogIn PID_read2(p16);       // control signal  from C&M device 2
AnalogIn PID_read3(p17);       // control signal  from C&M device 3

AnalogOut Control(p18);       // control signal connected to input of plant

void send_data(void);         // function that sends signal to plant 

float Kp;                     // float variable that will have value of received Kp 
float Ki;                     // float variable that will have value of received Ki
float Kd;                     // float variable that will have value of received Kd

C12832 lcd(p5, p7, p6, p8, p11);   // lcd pins on application board

I2CSlave slave(p28, p27);           //Configure I2C Slave

union Float                         // to convert floats to 4 bytes
{ 
    float m_float;
    uint8_t m_bytes[sizeof(float)];
};

uint8_t      bytes[sizeof(float)];

int main()
{ 
mode_s1=1;                              // select control mode for C&M device 1
mode_s2=0;                              // select monitoring mode for C&M device 2
mode_s3=0;                              // select monitoring mode for C&M device 3

slave.address(0x52);                    // address of C&M device 1
readdata.attach(&send_data,period);     // sending control signal to plant input 
while(1)
    {
    ledon=1;                    // shows output signal coordinator is on
    ledrun=!ledrun;   // blinking led to show output signal coordinator is running
    char rcd[5];         // array of 5 that stores bytes received
    char *rcdPTR;
    rcdPTR = &rcd[0];
    slave.receive();            // starts receiving data
    slave.read(rcdPTR,5);       // read data and store it in array of 5
   
 bytes[0] = rcd[1];             // first byte received stored in rcd[1]
 bytes[1] = rcd[2];             // second byte received stored in rcd[2]
 bytes[2] = rcd[3];             // third byte received stored in rcd[3]
 bytes[3] = rcd[4];             // fourth byte received stored in rcd[4]
           
  


if (rcd[0] == SetKp)  /* fifth byte stored in rcd[0] will be register in which gain                           
                      values are stored so if it has SetKp value then its value of     
                                                        proportional gain.*/

{
            Kp= *(float*)(bytes);
 }
 
 if (rcd[0] == SetKi)
 
{
            Ki= *(float*)(bytes);
 }
 
if (rcd[0] == SetKd)
 {
            Kd= *(float*)(bytes);
  }
  
wait (0.002);
  
  lcd.locate(0,8);                     // display value of kp on lcd
  lcd.printf("P=%2.2f%",Kp);
     
  lcd.locate(40,8);
  lcd.printf("I=%2.2f%",Ki);          // display value of ki on lcd
        
  lcd.locate(80,8);
  lcd.printf("D=%2.2f%",Kd);         // display value of kd on lcd

  
   wait(0.002);
         }
 }

void send_data(void)
{ 
if (mode_s1==1)//C&M device 1 is in control mode so its signal is transferred to plant
Control=PID_read1;
if (mode_s2==1)// C&M device 2 is in control mode so its signal is transferred to plant            
Control=PID_read2;
if (mode_s3==1)// C&M device 3 is in control mode so its signal is transferred to plant
Control=PID_read3;     
}


