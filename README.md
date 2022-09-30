# Client_PING
C#/WPF client useful to keep track of network devices, if they are online or not

## Release

### v1.0

### Manuale
TODO

#Example of configuration:
## Homepage

This application is useful to keep track of devices in your network at work or at home. You can import or insert a list of devices, keeping track of their pings.
Tha table sows the status of the ping and the last successful response received from the device. When you select a device, with the buttons below you can 
launch your specific application on the target device, for example an ssh/ftp/scp/... client, open the device in the browser or simply start a ping -t prompt.
It becomes very handy when you have to keep track of a list of PLCs at work or check the online status of different gateways in your LAN.

![alt text](https://github.com/Fedex1515/Client_PING/blob/master/Client_PING/Screenshots/Tab_1_Home.PNG?raw=true)

##Configuration tab:

The following image shows an example of configuration with different applications and required parameters passed to them. 
You can launch your specific application passing different arguments, for example the ip address of the device or username/password for the authentication process:

![alt text](https://github.com/Fedex1515/Client_PING/blob/master/Client_PING/Screenshots/Tab_3_Settings.PNG?raw=true)

##Device tab:

 The device tab shows a summary of the selected device:
 
 ![alt text](https://github.com/Fedex1515/Client_PING/blob/master/Client_PING/Screenshots/Tab_2_Device.PNG?raw=true)
