#include "mbed.h"
#define period 0.1              //define sampling time 
#include "C12832.h"

C12832 lcd(p5,p7,p6,p8,p11);   // pins to connect to application board LCD

const int addr = 0x52;         // address of C&M device 1

I2C i2c_master (p28, p27);      // configure i2c pins

void led_on (void);             // function to turn the LED on to show C&M is on
void signal_out (void);            //function to run device in control mode
void set_gains  (void);            //function to set gains after switching control signal off
void send_gains(void);             // function to send gains via i2c
void monitoring_out(void);         // function to run device in monitoring mode

float plant_integral = 0;         // initial value of integral
float previous_error = 0;         // initial value of error

AnalogIn Plant_read(p20);       //Analog output of plant
AnalogIn Ref_read(p19);         //Setpoint from signal generator
AnalogIn P_read(p15);           //proportional gain from potentiometer
AnalogIn I_read(p16);           //integral gain from potentiometer
AnalogIn D_read(p17);           //derivative gain from potentiometer

AnalogOut ControlOut(p18);    // control signal that goes to output signal coordinator

DigitalIn signal(p21);       // signal from switch to turn control signal off to change gain
DigitalIn mode(p22);         // mode selection signal from output signal coordinator

Ticker cmled;                     // set ticker for device on led
Ticker Out;                       //set ticker to generate control signal
Ticker stop_out;                  // ticker to stop the control signal and set gains again
Ticker send_data;                 // ticker to send gain values
Ticker monitoring_mode;           //ticker for monitoring mode

DigitalOut myled(LED1);          // LED to show device is on 
DigitalOut myled2(LED2);         // LED to show device is in control mode
DigitalOut myled3(LED3);         // LED to show device is in monitoring mode

float Kp;                                //float variable for proportional gain 
float Ki;                                //float variable for integral gain 
float Kd;                                //float variable for derivative gain 

union Float {                            // function to convert float into bytes
    float    m_float;
    uint8_t  m_bytes[sizeof(float)];
};
 
float        P_Gain = P_read*20;         // scaling potentiometer output for gains
float        I_Gain = I_read*20;
float        D_Gain = D_read*20;




uint8_t      bytes[sizeof(float)];
Float        myFloat;

void sendP_Gain(void)           // function to send proportional gain via i2c
{
//send a single byte of data, in correct I2C package
char cmd = 0x01;
*(float*)(bytes) = P_Gain;  // convert float to bytes
i2c_master.start(); //force a start condition
i2c_master.write(addr); //send the address
i2c_master.write(cmd); //send first byte of data, register
i2c_master.write(bytes[0]); //send  second byte of data, byte1 of float P_Gain
i2c_master.write(bytes[1]); //send third byte of data, byte 2 of float P_Gain
i2c_master.write(bytes[2]); //send fourth byte of data, byte 3 of float P_Gain
i2c_master.write(bytes[3]); //send fifth byte of data, byte 4 of float P_Gain
i2c_master.stop(); //force a stop condition
wait(0.002);
}

void sendI_Gain(void)
{
//send a single byte of data, in correct I2C package
char cmd = 0x02;
*(float*)(bytes) = I_Gain;  // convert float to bytes
i2c_master.start(); //force a start condition
i2c_master.write(addr); //send the address
i2c_master.write(cmd); //send first byte of data, register
i2c_master.write(bytes[0]); // send  second byte of data, byte1 of float I_Gain
i2c_master.write(bytes[1]); // send  third byte of data, byte2 of float I_Gain
i2c_master.write(bytes[2]); // send fourth byte of data, byte3 of float I_Gain
i2c_master.write(bytes[3]); // send fifth byte of data, byte4 of float I_Gain
i2c_master.stop(); //force a stop condition
wait(0.002);
}

void sendD_Gain(void)
{
//send a single byte of data, in correct I2C package
char cmd = 0x03;
*(float*)(bytes) = D_Gain;  // convert float to bytes
i2c_master.start(); //force a start condition
i2c_master.write(addr); //send the address
i2c_master.write(cmd); // send first byte of data, register
i2c_master.write(bytes[0]); // send  second byte of data, byte1 of float D_Gain
i2c_master.write(bytes[1]); // send  third byte of data, byte2 of float D_Gain
i2c_master.write(bytes[2]); // send fourth byte of data, byte3 of float D_Gain
i2c_master.write(bytes[3]); // send fifth byte of data, byte4 of float D_Gain
i2c_master.stop(); //force a stop condition
wait(0.002);
}

int main() {
    cmled.attach(&led_on,period);         // led1 is turned on to show that device is on
    if (mode==1)                          //to run in control mode
 {
      i2c_master.frequency(10000);
     send_data.attach(&send_gains,period); //function to send values of gains every 0.1 sec
     lcd.cls();                            // clear lcd
    lcd.locate(0,19);                                            
    lcd.printf("PID O/P:");
   if (signal==0)                          // switch to control mode 
   {
    lcd.locate(0,10);
    lcd.printf("Plant O/P:");
    Out.attach(&signal_out,period);        // function generate control signal
                  
        myled2=1;                          // led that shows device is in control mode
         }
    else
    {  
        myled2=0;                  // led turn off to show control output is zero 
    stop_out.attach(&set_gains,period);// function to set gains after stopping control signal
      }
}
if (mode==0)                         // device runs in monitoring mode
    {
      myled3=1;
      lcd.cls();
     lcd.locate(0,10);
    lcd.printf("Plant O/P:");
    monitoring_mode.attach(&monitoring_out,period);   // function to monitor plant output

     }
}
void led_on(void)  // function to turn led 1 on
{
    myled =1;
}
void signal_out (void)       // function to run device in control mode

{
    Kp=P_read*20;
    Ki=I_read*20;
    Kd=D_read*20;

    float plant_sample = Plant_read; //Read plant output
    float ref_sample = Ref_read;    //Read setpoint
    
    float plant_error = ref_sample - plant_sample; // calculate error
   plant_integral = plant_integral + 
    ((plant_error+previous_error)*period/2); //find integral using area of  the trapezium
    
  float plant_derivative = (plant_error - previous_error)/period;  //find derivative
    
    float output = (plant_error*Kp)+(plant_integral*Ki)+(plant_deriviative*Kd)/3.3;
      
    if (output >1) output = 1;
    if (output <0) output = 0;
    
    previous_error = plant_error;
    ControlOut = output;
    lcd.locate(0,1);
    lcd.printf("kp=%2.2f,ki=%2.2f,kd=%2.2f", Kp,Ki,Kd); // display gain values on lcd
    lcd.locate(45,10);
    lcd.printf("%1.2f", plant_sample*3.3);             // display plant output on lcd
    lcd.locate(45,19);
     lcd.printf("%1.2f",output*3.3);                  // display control signal value on lcd
     wait(0.05); 
    }
    
    
void set_gains (void)           // function to turn control signal off and set gains
{
 ControlOut=0;
Kp=P_read*20;
Ki=I_read*20;
Kd=D_read*20;

    
   lcd.locate(0,1);
   lcd.printf("Set the Gains");
   lcd.locate(0,10);
   lcd.printf("kp=%2.2f,ki=%2.2f,kd=%2.2f", Kp,Ki,Kd);  // display new gain values
   lcd.locate(45,19);
    lcd.printf("%1.2f",ControlOut*3.3);
     wait(0.05);
   
}
    
   
    void send_gains(void)         // function that send gain values
    {
    sendP_Gain();
    sendI_Gain();
    sendD_Gain();
     
     }
void monitoring_out (void)      // function to run device in monitoring mode
{
    float plant_sample = Plant_read;
    lcd.locate(45,10);
     lcd.printf("%1.2f", plant_sample*3.3);   // display value of plant output
       wait(0.05);
    }









