import { useEffect, useState } from "react";
import { useOfficeContext } from "../OfficeContext";

const SideBar = ({officeState}) => {
  const { state } = useOfficeContext();

  const [emailBody, setEmailBody] = useState('');
  const [isComposeMode, setIsComposeMode] = useState(false)


  useEffect(() => {
    window.Office?.onReady((info) => {
      setEmailBody(window.Office.context.mailbox.item);

      const item = window.Office.context.mailbox.item

      if(item.body){
        item.body.getAsync(window.Office.CoercionType.Text, (result) => {
          setEmailBody(result.value)
        })
      }

      if(item.itemType === window.Office.MailboxEnums.ItemType.Message){
        setIsComposeMode(true)
      }
    })
  }, [window.Office])

  const handleAddContent = () => {
        // Get the current appointment item (email in compose mode)
        const item = window.Office.context.mailbox.item;
  
        // Check if the item has a body
        if (item.body) {
          // Text to append to the email body
          const textToAppend = "\n\n----\nThis is additional text added by the add-in.";
  
          // Append the text to the existing email body
          item.body.getAsync(window.Office.CoercionType.Text, (result) => {
            if (result.status === window.Office.AsyncResultStatus.Succeeded) {
              const currentBody = result.value;
              const updatedBody = currentBody + textToAppend;
  
              // Set the updated body with the appended text
              item.body.setAsync(updatedBody, { coercionType: window.Office.CoercionType.Text }, (result) => {
                if (result.status === window.Office.AsyncResultStatus.Succeeded) {
                  console.log('Text appended to the email body successfully.');
                } else {
                  console.error('Error setting email body:', result.error.message);
                }
              });
            } else {
              console.error('Error getting email body:', result.error.message);
            }
          });
    }
  }

  const handleAddSubject = () => {
    // Get the current appointment item (email in compose mode)
    const item = window.Office.context.mailbox.item;
  
    // Check if the item has a body
    if (item.body) {
      // Text to append to the email body
      const textToAppend = "This is the new subject " + new Date();

      // Append the text to the existing email body
      item.subject.setAsync(textToAppend, { coercionType: window.Office.CoercionType.Text }, (result) => {
        if (result.status === window.Office.AsyncResultStatus.Succeeded) {
          console.log('Text appended to the email body successfully.');
        } else {
          console.error('Error setting email body:', result.error.message);
        }
      });
    }
  }
  

  return <>
    {isComposeMode && <button onClick={handleAddContent}>Add to email body</button>}
    {isComposeMode && <button onClick={handleAddSubject}>Add to email subject</button>}
     <h4>{JSON.stringify(emailBody)}</h4>
     </>;
}

export default SideBar