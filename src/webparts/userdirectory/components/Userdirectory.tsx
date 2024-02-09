//ORIGINAL CODE WORKS 
// import * as React from 'react';
// import {Card } from 'antd';
// import 'antd/dist/reset.css';
// import { useEffect, useState } from "react";
// // import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
// import { MSGraphClientV3 } from '@microsoft/sp-http';
// import { IUserdirectoryProps } from './IUserdirectoryProps';

// export default function Userdirectory(props:IUserdirectoryProps){
//   const [users, setUsers] = useState([]);
//   const [searchQuery, setSearchQuery] = useState("");
//     const copyToClipboard = (email: any) => {
//     navigator.clipboard.writeText(email);
//   };

//   useEffect(()=>{
//     props.context.msGraphClientFactory
//     .getClient("3")
//     .then((client:MSGraphClientV3) =>{
//       client
//       .api('/users')
//       .version("v1.0")
//       .select("id,displayName,department,jobTitle,mail,userPrincipalName")
//       .get((error: any, eventsResponse, rawResponse?: any) => {
//         if (error) {
//           console.error("Message is: " + error);
//           return;
//         }

//         const userData = eventsResponse.value;
//         console.log(userData);
//         setUsers(userData);
//     });
//   });
//   },[props.context.msGraphClientFactory]);
//   console.log(users);
//   const filteredUsers = users.filter((user: any) =>
//     user.displayName.toLowerCase().includes(searchQuery.toLowerCase())
//   );
  
//   return(
// <>
// <div>
//   <div style={{display:"flex"}}>
//     <div style={{borderLeft:"5px solid #018FD4", borderRadius:"5px", width:"50%", }}>
//       <span style={{lineHeight:"3.5"}}>Employee Directory</span>
//       </div>
//     <div style={{width:"48%", textAlign:"end", marginTop:"15px"}}>
//       <img src={require("../assets/search.svg")} style={{position:"relative", left:"20px"}}/>
//       <input type="text"  value={searchQuery} onChange={(e)=>{
//         setSearchQuery(e.target.value)
//       }}
//       placeholder='Search by name' style={{padding:"5px 10px 5px 25px"}}/>
//     </div>
//   </div>
// <div style={{display:"flex", justifyContent:"space-between", flexWrap:"wrap", overflowY:"scroll", height:"400px", marginTop:"20px"}}> 
// {filteredUsers.map((user:any) => (
//         <div key={user.id} style={{flexBasis:"48%", marginTop:"35px"}}>
//           <Card style={{ width: "100%", margin: 'auto', boxShadow: '0 6px 8px rgba(0, 0, 0, 0.1)'}}>
//             <div style={{display:"flex", gap:"5%"}}>
//           <div>
//           <img src={`/_layouts/15/userphoto.aspx?size=L&username=${user.userPrincipalName}`} alt={`${user.displayName}`}  style={{width:'50px', height:'50px', borderRadius:'50%'}}/>
//           </div>
//           <div>
//           <h3 style={{fontSize:'18px', fontWeight:'400', color:'#4D4D4D'}}>{user.displayName}</h3>
//           <h4 style={{fontSize:'12px', fontWeight:'350', color:'#9E9E9E'}}>{user.jobTitle}-{user.department}</h4>
//           </div>
//           </div>
//              <div style={{fontSize:'12px'}}>
//             <span style={{color:"#242424"}}>Mail:</span><span style={{color:"#018FD4"}}> {user.mail}</span>
//               <img
//                       src={require("../assets/copymail.svg")}
//                       alt="Copy"
//                       onClick={() => copyToClipboard(user.mail)}
//                       style={{
//                         cursor: "pointer",
//                         width: "12px",
//                         marginLeft: "10px",
//                       }} 
//                       />
//           </div>
//           </Card>
//         </div>
//       ))}
// </div>
// </div>
// </>  
// );
// }


import * as React from 'react';
import { Card } from 'antd';
import 'antd/dist/reset.css';
import { useEffect, useState } from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IUserdirectoryProps } from './IUserdirectoryProps';
import { Requestmail } from '../../MailTrigger/MailTrigger';

export default function Userdirectory(props: IUserdirectoryProps) {
  const [users, setUsers] = useState<any[]>([]);
  const [currentuser, setCurrentuser] = useState<any[]>([]);

  const [searchQuery, setSearchQuery] = useState('');
  const [userMessages, setUserMessages] = useState<{ [key: string]: string }>({});

  let senderName: string = currentuser[0]?.displayName || '';
  let senderJob: string = currentuser[0]?.jobTitle || '';
  let senderDept: string = currentuser[0]?.department || '';

   console.log(senderName);
   console.log(senderJob);
   console.log(senderDept);
   

  const copyToClipboard = (email: any) => {
    navigator.clipboard.writeText(email);
  };

  const handleSendMessage = (userMail: string, displayName:string, senderName:string, senderJob: string, senderDept: string) => {
    const msg = userMessages[userMail];
    if (!msg || !msg.trim()) {
      console.error('Message cannot be empty');
      return;
    }

    Requestmail(msg,userMail,displayName,senderName,senderJob,senderDept );
    setUserMessages({ ...userMessages, [userMail]: '' });
  };


  console.log(userMessages);


  useEffect(() => {
    props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        client
          .api('/users')
          .version('v1.0')
          .select('id,displayName,department,jobTitle,mail,userPrincipalName')
          .get((error: any, eventsResponse: { value: any[] }, rawResponse?: any) => {
            if (error) {
              console.error('Message is: ' + error);
              return;
            }

            const userData = eventsResponse.value;
            // console.log(userData);
            setUsers(userData);
          });
      });
  }, [props.context.msGraphClientFactory]);

  // console.log(users);


  
  useEffect(() => {
    props.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        client
          .api('/me')
          .version('v1.0')
          .select('jobTitle,department,displayName,mail')
          .get((error: any, eventsResponse: any, rawResponse?: any) => {
            if (error) {
              console.error('Error fetching current user data:', error);
              return;
            }
  
            console.log('Raw response:', rawResponse); // Log raw response for debugging
            console.log('Current User Data:', eventsResponse);
            setCurrentuser([eventsResponse]);
          });
      });
  }, [props.context.msGraphClientFactory]);
  
  
  
  console.log(currentuser);

  useEffect(() => {
    // Monitor changes in currentuser and update sender information
    const [firstUser = {}] = currentuser; // Destructure the first item with default empty object
    const { displayName: senderName, jobTitle: senderJob, department: senderDept } = firstUser;
  
    console.log(senderName);
    console.log(senderJob);
    console.log(senderDept);
  }, [currentuser]);
  // const sender:string=currentuser.displayName

  const filteredUsers = users.filter((user: any) =>
    user.displayName.toLowerCase().includes(searchQuery.toLowerCase())
  );
  console.log(props.context.pageContext.user)

  return (
    <>
      <div>
        <div style={{ display: 'flex' }}>
          <div style={{ borderLeft: '5px solid #018FD4', borderRadius: '5px', width: '50%' }}>
            <span style={{ lineHeight: '3.5', fontSize:"18px", fontWeight:"700" }}>Employee Directory</span>
          </div>
          <div style={{ width: '48%', textAlign: 'end', marginTop: '15px' }}>
            <img src={require('../assets/search.svg')} style={{ position: 'relative', left: '20px', top:"-1px" }} />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
              }}
              placeholder="Search by name"
              style={{ padding: '5px 10px 5px 25px' }}
            />
          </div>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', overflowY: 'scroll', height: '400px', marginTop: '20px' }}>
          {filteredUsers.map((user: any) => (
            <div key={user.id} style={{ flexBasis: '48%', marginTop: '35px' }}>
              <Card style={{ width: '100%', margin: 'auto', boxShadow: '0 6px 8px rgba(0, 0, 0, 0.1)' }}>
                <div style={{ display: 'flex', gap: '5%' }}>
                  <div>
                    <img
                      src={`/_layouts/15/userphoto.aspx?size=L&username=${user.userPrincipalName}`}
                      alt={`${user.displayName}`}
                      style={{ width: '50px', height: '50px', borderRadius: '50%' }}
                    />
                  </div>
                  <div>
                    <h3 style={{ fontSize: '18px', fontWeight: '400', color: '#4D4D4D' }}>{user.displayName}</h3>
                    <h4 style={{ fontSize: '12px', fontWeight: '350', color: '#9E9E9E' }}>{user.jobTitle}-{user.department}</h4>
                  </div>
                </div>
                <div style={{ fontSize: '12px' }}>
                  <span style={{ color: '#242424' }}>Mail:</span><span style={{ color: '#018FD4' }}> {user.mail}</span>
                  <img
                    src={require('../assets/copymail.svg')}
                    alt="Copy"
                    onClick={() => copyToClipboard(user.mail)}
                    style={{
                      cursor: 'pointer',
                      width: '12px',
                      marginLeft: '10px',
                    }}
                  />
                </div>
                <div>
                  <input
                    type="text"
                    value={userMessages[user.mail] || ''}
                    onChange={(e) => {
                      setUserMessages({ ...userMessages, [user.mail]: e.target.value });
                    }}
                  />
                  <button onClick={() => handleSendMessage(user.mail, user.displayName,senderName,senderJob,senderDept)}>Send</button>
                </div>
              </Card>
            </div>
          ))}
        </div>
      </div>
    </>
  );
}
