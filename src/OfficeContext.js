import React, { useState, useEffect, useContext } from "react";


const OfficeContext = React.createContext({});

export const useOfficeContext = () => useContext(OfficeContext);

export const OfficeContextProvider = ({ children }) => {
  const [state, setState] = useState({});

  useEffect(() => {
    const asyncHandler = async () => {
      // We'll call this function to read the current office state
      // from mailbox, item.
      const update = () => {
        if(!window.Office) return;
        // Unpack the office context.
        const mailbox = window.Office.context.mailbox;
        const item = mailbox?.item;
        const getItemState = () => {
          if (!item) return undefined;
          const {
            from,
            to,
            internetMessageId,
            conversationId,
            itemType,
            cc,
            body,
          } = item;
          return {
            from,
            to,
            cc,
            internetMessageId,
            conversationId,
            itemType,
            body,
          };
        };

        const getMailboxState = () => {
          if (!mailbox) return undefined;

          const { userProfile } = mailbox;

          return {
            userProfile,
            item: getItemState(),
          };
        };

        setState({
          state: {
            mailbox: getMailboxState(),
          },
          isInitialized: true,
        });
      };

      // Update our application, now that Office is initialized.
      update();
    };
    window.Office.onReady(() => {
      asyncHandler();
    });
  }, []);

  return (
    <OfficeContext.Provider value={state}>{children}</OfficeContext.Provider>
  );
};

export const OfficeContextConsumer = OfficeContext.Consumer;