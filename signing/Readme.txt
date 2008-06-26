This folder contains the necessary scripts to create a signed build of the OdfConverter solution. 

The certificate provided here is a self-signed certificated for demonstration purpose only.

To enable code signing, please do the following:

- create your own .snk file using the following command line from the Visual Studio command prompt
    sn -k OdfConverter.snk

- extract the public key using the following command
    sn -p OdfConverter.snk OdfConverter.public.snk

- display the public key token using 
    sn -tp OdfConverter.public.snk    
  and enter the displayed value in the file config.h in the COM shim projects

- rename the file sign.sample.bat to sign.bat and configure your certificate and certificate password accordingly