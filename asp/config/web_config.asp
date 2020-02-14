<%
'###################################################################################################
'#### PARAMETERS for QuickerSite - Check http://www.quickersite.com/support                     ####
'###################################################################################################

'###################################################################################################
'#### Use ONLY if you install QuickerSite in a VIRTUAL DIRECTORY in IIS            
'###################################################################################################
const C_VIRT_DIR=""	'virtual directory -> DO NOT FORGET trailing slash: eg: '/QuickerSite'
const C_DIRECTORY_QUICKERSITE=""  'COPY the value of C_VIRT_DIR here ! 

'###################################################################################################
'#### Folder that stores USERFILES (uploaded files & images) 
'#### The IUSR needs read & write permissions on this folder!     
'###################################################################################################
Application("QS_CMS_userfiles")="/userfiles/" 'do not forget both slashes, at the start & end!

'###################################################################################################
'#### MAILER COMPONONENT - SUPPORTED ASP COMPONENTS: persits.mailsender;cdo.message;jmail.message
'#### cdonts.newmail, smtpsvg.mailer 
'###################################################################################################
const C_MAILCOMPONENT="cdo.message"
const C_SMTPSERVER="localhost" ' your own smtp server
const C_SMTPPORT=25 'port smtp server
const C_SMTPUSERNAME="" 'smtp user - optional - only if needed!
const C_SMTPUSERPW="" 'smtp user password - optional - only if needed!
QSCDO_smtpusessl=false 'use SSL for SMTP server? Needed (set to true!) for smtp.gmail.com

'###################################################################################################
'#### C_SENDUSING: applies to cdo.message only
'#### 1 (default) for local smtp-service (eg: "localhost" or "127.0.0.1")
'#### 2 for external smtp-service (eg: mail.mydomain.com)
'#### 3 for Exchange server
const C_SENDUSING=1
'###################################################################################################

'###################################################################################################
'#### PATH to databases (IUSR needs write-permissions)
'###################################################################################################
QS_DBS				= 1	'1 for Access
C_DATABASE			= server.MapPath (C_VIRT_DIR & "/db/QuickerSite.mdb") 'RENAME THIS FILE!
'C_DATABASE			= "D:\inetpub\wwwroot\db\mydb.mdb" 'you can also hardcode the path to the db!
C_DATABASE_LABELS	= server.MapPath (C_VIRT_DIR & "/db/QuickerLabels.mdb")

'###################################################################################################
'#### MS SQLSERVER CONNECTION SETTINGS - leave all parameters BLANK in case you use ACCESS!
'#### If you use SQLServer, you still HAVE to (re)move/name "/db/data_dp0ct4rx.mdb" 
'#### If you use SQLServer, you still NEED "/db/labels_ue9yy3pg.mdb"! Do not comment that line above!
'#### UNCOMMENT THE LINES BELOW IN CASE YOU WANT TO USE SQL SERVER!!!
'###################################################################################################
'QS_DBS			= 2					'2 for SQL Server 2000/2005
'SQL2005_SERVER	= ""				'server
'SQL2005_DB		= ""				'database
'SQL2005_UID	= ""				'user
'SQL2005_PWD	= ""				'password

'###################################################################################################
'#### This value corresponds to the value for iId in the table tblCustomer of your QuickerSite DB
'#### You can have multiple sites in 1 database and running on 1 codebase. You then have to move
'#### this variable to the global.asa of every QuickerSite you host and comment this line!
'###################################################################################################
Application("QS_CMS_iCustomerID")=73

'###################################################################################################
'#### THESE 3 PARAMS ARE THERE FOR SPECIFIC PURPOSES ONLY
'###################################################################################################
const C_ADMINEMAIL="" 'the email address to forward VBScript runtime errors to
const C_ADMINPASSWORD="" 'you must set this password to get access to the admin-section
const QS_admin_login_page="ad_login.asp" 'asp page that logs on to the admin-part

'###################################################################################################
'#### BELOW YOU CAN PROVIDE 2 CUSTOM ASP PAGES THAT WILL BE EXECUTED AT EVERY PAGELOAD
'###################################################################################################
'The paths below must be relative to the root of QuickerSite - RENAME FILES FOR SECURITY-REASONS
execBeforePageLoad	= ""'/common/before.asp"' eg: "common/before.asp" 'asp to be executed BEFORE any (public) pageload
execAfterPageLoad	= ""' eg: "common/after.asp" 'asp to be executed AFTER any (public) pageload

'###################################################################################################
'#### Miscellaneous settings
'###################################################################################################

const QS_backsite_login_page="bs_login.asp" 'asp page that logs on to the backsite
const QS_number_of_allowed_attempts_to_login=5 'number of attempts to logon to the backsite/admin
manyContacts=250 'when should the number of contacts be considered "many"
const C_DEV=false 'True takes your QuickerSite(s) offline (for upgrading purposes)

sNewTemplatesPath=server.MapPath (C_VIRT_DIR & "/templates")
sNewTemplatesURL=C_VIRT_DIR & "/templates"
bBrowseOnlineTemplates=true 'enables the download of new templates
sAffArtisteer=""
%> 






















