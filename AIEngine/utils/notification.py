import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from utils import logging

def send(base, send_to, subject, content, file=None):
  smtp_server = "mxa-00181c01.gslb.pphosted.com"
  port = 25  # For starttls
  sender_email = "donotreply@asia-assistance.com"
  #password = input("Type your password and press enter: ")

  #convert string for receiver email to lowercase
  receiver_email = send_to.lower()
  # Create a secure SSL context
  context = ssl.create_default_context()

  message = MIMEMultipart("alternative")
  message["Subject"] = subject
  message["From"] = sender_email
  message["To"] = receiver_email

  # Try to log in to server and send email
  try:
      server = smtplib.SMTP(smtp_server,port)
      server.ehlo() # Can be omitted
      server.starttls(context=context) # Secure the connection
      server.ehlo() # Can be omitted
      #server.login(sender_email, password)
      html = """\
                <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
                <html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
                <head>
	                <!--[if gte mso 9]>
	                <xml>
		                <o:OfficeDocumentSettings>
		                <o:AllowPNG/>
		                <o:PixelsPerInch>96</o:PixelsPerInch>
		                </o:OfficeDocumentSettings>
	                </xml>
	                <![endif]-->
	                <meta http-equiv="Content-type" content="text/html; charset=utf-8" />
	                <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1" />
                    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
	                <meta name="format-detection" content="date=no" />
	                <meta name="format-detection" content="address=no" />
	                <meta name="format-detection" content="telephone=no" />
	                <meta name="x-apple-disable-message-reformatting" />
                    <!--[if !mso]><!-->
	                <link href="https://fonts.googleapis.com/css?family=Muli:400,400i,700,700i" rel="stylesheet" />
                    <!--<![endif]-->
	                <title>Email Template</title>
	                <!--[if gte mso 9]>
	                <style type="text/css" media="all">
		                sup { font-size: 100% !important; }
	                </style>
	                <![endif]-->
	

	                <style type="text/css" media="screen">
		                /* Linked Styles */
		                body { padding:0 !important; margin:0 !important; display:block !important; min-width:100% !important; width:100% !important; background:#001736; -webkit-text-size-adjust:none }
		                a { color:#66c7ff; text-decoration:none }
		                p { padding:0 !important; margin:0 !important } 
		                img { -ms-interpolation-mode: bicubic; /* Allow smoother rendering of resized image in Internet Explorer */ }
		                .mcnPreviewText { display: none !important; }

				
		                /* Mobile styles */
		                @media only screen and (max-device-width: 480px), only screen and (max-width: 480px) {
			                .mobile-shell { width: 100% !important; min-width: 100% !important; }
			
			                .text-header,
			                .m-center { text-align: center !important; }
			
			                .center { margin: 0 auto !important; }
			                .container { padding: 20px 10px !important }
			
			                .td { width: 100% !important; min-width: 100% !important; }

			                .m-br-15 { height: 15px !important; }
			                .p30-15 { padding: 30px 15px !important; }

			                .m-td,
			                .m-hide { display: none !important; width: 0 !important; height: 0 !important; font-size: 0 !important; line-height: 0 !important; min-height: 0 !important; }

			                .m-block { display: block !important; }

			                .fluid-img img { width: 100% !important; max-width: 100% !important; height: auto !important; }

			                .column,
			                .column-top,
			                .column-empty,
			                .column-empty2,
			                .column-dir-top { float: left !important; width: 100% !important; display: block !important; }

			                .column-empty { padding-bottom: 10px !important; }
			                .column-empty2 { padding-bottom: 30px !important; }

			                .content-spacing { width: 15px !important; }
		                }
	                </style>
                </head>
                <body class="body" style="padding:0 !important; margin:0 !important; display:block !important; min-width:100% !important; width:100% !important; background:#001736; -webkit-text-size-adjust:none;">
	                <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#001736">
		                <tr>
			                <td align="center" valign="top">
				                <table width="650" border="0" cellspacing="0" cellpadding="0" class="mobile-shell">
					                <tr>
						                <td class="td container" style="width:650px; min-width:650px; font-size:0pt; line-height:0pt; margin:0; font-weight:normal; padding:55px 0px;">
							                <!-- Header -->
							                
							                <!-- END Header -->

							                <repeater>
								                <!-- Intro -->
								                <layout label='Intro'>
									                <table width="100%" border="0" cellspacing="0" cellpadding="0">
										                <tr>
											                <td style="padding-bottom: 10px;">
												                <table width="100%" border="0" cellspacing="0" cellpadding="0">
													                <tr>
														                <td class="tbrr p30-15" style="padding: 60px 30px; border-radius:26px 26px 0px 0px;" bgcolor="#12325c">
															                <table width="100%" border="0" cellspacing="0" cellpadding="0">
																                <tr>
																	                <td class="h1 pb25" style="color:#ffffff; font-family:'Muli', Arial,sans-serif; font-size:40px; line-height:46px; text-align:center; padding-bottom:25px;"><multiline>Asia Assistance - Robotic Process Automation</multiline></td>
																                </tr>
																                <tr>
																	                <td class="text-center pb25" style="color:#c1cddc; font-family:'Muli', Arial,sans-serif; font-size:16px; line-height:30px; text-align:center; padding-bottom:25px;"><multiline>"""+content+"""</multiline></td>
																                </tr>
																                <!-- Button -->
																                <tr>
																	                <td align="center">
																		                
																	                </td>
																                </tr>
																                <!-- END Button -->
															                </table>
														                </td>
													                </tr>
												                </table>
											                </td>
										                </tr>
									                </table>
								                </layout>
								                <!-- END Intro -->
							                <!-- Footer -->
							                <table width="100%" border="0" cellspacing="0" cellpadding="0">
								                <tr>
									                <td class="p30-15 bbrr" style="padding: 50px 30px; border-radius:0px 0px 26px 26px;" bgcolor="#0e264b">
										                <table width="100%" border="0" cellspacing="0" cellpadding="0">
											                <tr>
												                <td align="center" style="padding-bottom: 30px;">
													                Power by Data & Technology Innovation &copy; 
												                </td>
											                </tr>
											                <tr>
												                <td class="text-footer1 pb10" style="color:#c1cddc; font-family:'Muli', Arial,sans-serif; font-size:16px; line-height:20px; text-align:center; padding-bottom:10px;"><multiline>Robotic Process Automation - RPA</multiline></td>
											                </tr>
											                <tr>
												                <td class="text-footer2" style="color:#8297b3; font-family:'Muli', Arial,sans-serif; font-size:12px; line-height:26px; text-align:center;"><multiline>
                                        Asia Assistance<br />
                                        AA One, Block N, Jaya One<br />
                                        72A Jalan Universiti<br />
                                        46200 Petaling Jaya<br />
                                        Selangor Darul Ehsan, Malaysia<br />
                                        </multiline></td>
											                </tr>
										                </table>
									                </td>
								                </tr>
								               
							                </table>
							                <!-- END Footer -->
						                </td>
					                </tr>
				                </table>
			                </td>
		                </tr>
	                </table>
                </body>
                </html>

             """
      #Turn these into plain/html MIMEText objects
      #htmlformat = MIMEText(text, "plain")
      htmlformat = MIMEText(html, "html")

      # Add HTML/plain-text parts to MIMEMultipart message
      # The email client will try to render the last part first
      
      message.attach(htmlformat)

      # Open PDF file in binary mode
      if file != None: 
        with open(file, "rb") as attachment:
          # Add file as application/octet-stream
          # Email client can usually download this automatically as attachment
          part = MIMEBase("application", "octet-stream")
          part.set_payload(attachment.read())

        #Encode file to ASCII characters
        encoders.encode_base64(part)
        # Add header as key/value pair to attachment part
        part.add_header("Content-Disposition", f"attachment; filename= {file}")
        message.attach(part)

      # TODO: Send email here
      server.sendmail(sender_email, receiver_email, message.as_string())
  except Exception as e:
      # Print any error messages to stdout
      print(e)
      logging('notification', 'email notification error: {0}'.format(e), base)
  finally:
      server.quit()


#Sample code for trigger notification
#html = """\
 #   <p>Hi,<br>
  #     How are you?<br>
   #    <a href="http://www.realpython.com">Real Python</a> 
    #   has many great tutorials.
    #</p>
#"""
#send("eikden.yeoh@asia-assistance.com","RPA Notification", html)
