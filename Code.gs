//*************************************  This is the start of Variables to set for Application to run   ****************************
// #1      **************** Type of files to run.
var fileTypesToExtract = ['jpg', 'tif', 'png', 'gif', 'bmp', 'svg' , 'xls' , 'xlsx' , 'txt' , 'docx' , 'doc' , 'rtf' , 'jpeg'];
// #2      **************** This is the label to mark the email with after the attachment has been processed
var labelName = 'AlreadyProcessed';
// #3      **************** This is the name of the folder that will be created that the attachment will go in.
var folderToProcess = 'jnewfolder';
// #4      **************** This is the number of emails to search through
var num_of_emails_to_read = "99";
// #5      ****************  Put the subject of the email your searching for in this variable
var SUBJECT_to_search_for = "requisition";

//*************************************  This is the end of Variables to set for Application to run   ****************************

var row = -1;
var subject = "";
var from = "";
var Should_i_ProcessAnotherOne = true;


var is_this_the_subject_im_searchN4 = 0;   
var label ;
    

function doGet() {
    is_There_An_Email_To_Process();
      }
   
      
   function deletePreviousFolder() 
{
  
  var folderToDelete = '';  

  var folders = DriveApp.getFolders();
  while (folders.hasNext()) 
         {
    var folder2 = folders.next();      
     folderToDelete = folder2.getName();
     if(folderToDelete == folderToProcess){
     folder2.setTrashed(true);
     };
         }
};

function ProcessEmail(input)
{
   
deletePreviousFolder();
         
   SUBJECT_to_search_for = SUBJECT_to_search_for.toLowerCase(); 
      
   var threads = GmailApp.getInboxThreads(0,num_of_emails_to_read);
   var label = getLabelForGmail_(labelName);
   
  
          for (var col = 0; col < threads.length; col++)
       {
            var AllEmailMessages = threads[col].getMessages();
                   for(var TheEmailMsg in AllEmailMessages)
                   {
                    
                  row = row + 1
                
                
                
                  subject = AllEmailMessages[TheEmailMsg].getSubject();  
                  subject = subject.toLowerCase(); 
                       
                  from = AllEmailMessages[TheEmailMsg].getFrom();
                  from =  from.toLowerCase(); 
                                        
                     
                 
                             
                          
                            
                   is_this_the_subject_im_searchN4 = subject.indexOf(SUBJECT_to_search_for);
                           
                                                  
                        if (is_this_the_subject_im_searchN4 != -1)
                        {
                                                                   
                        var attachments = AllEmailMessages[TheEmailMsg].getAttachments();
                        
                          for(var k in attachments)
                                {
                                     var attachment = attachments[k];
                                     var isImageType = checkDocType_(attachment);
                                      if(!isImageType) continue;
                                      
                                 
                                     var attachmentBlob = attachment.copyBlob();
                                     var file = DriveApp.createFile(attachmentBlob);
                                     var newFileID = file.getId() ;
                                     var newFolder = DriveApp.createFolder(folderToProcess) ; 
                                 
                                            newFolder.createFile(file);
                                           
                                               DriveApp.getRootFolder().removeFile(file);
                                           
                                             threads[col].addLabel(label);
                                            threads[col].markRead()
                           }
                                                     
                        }
                                     
                        
                        if (row == num_of_emails_to_read)
                        {
                        return;
                        }
                                             
                   }
             }
          
             row++;
  }
  

function getLabelForGmail_(name){
  var label = GmailApp.getUserLabelByName(name);
  if(label == null){
 label = GmailApp.createLabel(name);
  }
  return label;
}


function checkDocType_(attachment){
  var fileName = attachment.getName();
  var temp = fileName.split('.');
  var fileExtension = temp[temp.length-1].toLowerCase();
  if(fileTypesToExtract.indexOf(fileExtension) != -1) return true;
 
  else return false;
}



function is_There_An_Email_To_Process()
{
            
   SUBJECT_to_search_for = SUBJECT_to_search_for.toLowerCase(); 
      
   var threads = GmailApp.getInboxThreads(0,num_of_emails_to_read);
   var label = getLabelForGmail_(labelName);
   var Found_One = false;
     
          for (var col = 0; col < threads.length; col++)
       {
         if(!Should_i_ProcessAnotherOne) continue;
            var AllEmailMessages = threads[col].getMessages();
                   for(var TheEmailMsg in AllEmailMessages)
                   {
                      if(!Should_i_ProcessAnotherOne) continue;
                  row = row + 1
                
                
                  subject = AllEmailMessages[TheEmailMsg].getSubject();  
                  subject = subject.toLowerCase(); 
                       
                  from = AllEmailMessages[TheEmailMsg].getFrom();
                  from =  from.toLowerCase(); 
                                        
                                
                            
                   is_this_the_subject_im_searchN4 = subject.indexOf(SUBJECT_to_search_for);
                           
                                                  
                        if (is_this_the_subject_im_searchN4 != -1)
                        {
                                                                   
                      if  (AllEmailMessages[TheEmailMsg].isUnread())
                                         {
                                          Found_One = true;
                                           ProcessEmail(Found_One)
                                             if(!Should_i_ProcessAnotherOne) continue;
                                           Should_i_ProcessAnotherOne = false;
                                         }
                        }
                   
                                                
                        
                        if (row == num_of_emails_to_read)
                        {
                        return;
                        }
                                             
                   }
             }
          
             row++;
  }
