import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'NavbarExtensionApplicationCustomizerStrings';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { AadTokenProvider, HttpClient, HttpClientResponse} from '@microsoft/sp-http';
import './assets/css/rpg_styles_nav.css';
import styles from './NavbarExtensionApplicationCustomizer.module.scss';
import * as $ from 'jquery';
import { app } from "@microsoft/teams-js";

const LOG_SOURCE: string = 'NavbarExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INavbarExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Top: string;
  Bottom: string;
  testMessage: string;
  description: string;
  navContents: string;
  MyAplication: string;
  dropBoxMyBook: string;
  MyBookmarks: string;
  notificationIconBox:string;
  navIcon: string;
  UserEmail: string;
  userEmail: string;
  pictureUrl: string;
  userName:string;
}

export interface ApplicationsListItem {
  description: string;
  seeAllButton: string;
  Title: string;
  URL: string;
  Company: string;
}

export interface NotificationsListItem {
  NotificationText: string;
  NotificationURL:{
    Url: string;
  }
  TargetAudience: string;
  Created: string;
  StartDate: string;
  Company: string;
}

export interface YammerNotifications{
  body:{
  plain: string;
  }  
  web_url: string;
  published_at: string;
}

export interface SharepointNotifications{
  createdDate: string;
  ID: string;
  Title: string;
  FileLeafRef: string;
  text: string;
  mentions:{
    length: number;
    email: string;
  }
}

export interface IVivaEngageWebPartProps {
  description: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class NavbarExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<INavbarExtensionApplicationCustomizerProperties> {

    private graphApiBookmarks: any[] = [];
    private notifications: any[] = [];
    private noOfNotifications: number = 0;
    private noOfApplications: number = 0;
    private noOfBookmarks: number = 0;
    private externalUser: boolean = false;
  
    private vivaEngageToken: string = '';
    private vivaEngagePosts: string = '';

    private users: any[] = [];
  
  // private domElement: HTMLElement | null = null; // Add this line to declare domElement
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  
  private async userInformation(): Promise<void> {
    this.properties.pictureUrl = require("./assets/iconPerson.png");
    this.properties.UserEmail = this.context.pageContext.user.email.toLowerCase();

   const currentUser = this.users.filter((user) => {
      return user.Email === this.properties.UserEmail;
   });

   console.log("Current User:", currentUser);

   if(currentUser.length == 1){
    this.properties.userName = currentUser[0].UserName;
    this.properties.pictureUrl = currentUser[0].ProfilePicture.Url;
   }else if(currentUser.length > 1){
    const multipleUsers = currentUser.filter((user) => {
      console.log("User",user.UserName," GroupCompany from List: ", user.GroupCompany);
      return user.GroupCompany;
    });
      console.log("Multiple Users:", multipleUsers);
      this.properties.userName = multipleUsers[0].UserName;
      if(multipleUsers[0].ProfilePicture){
        this.properties.pictureUrl = multipleUsers[0].ProfilePicture.Url;
      }
   }else{
    this.properties.userName = "Welcome, User";
    this.properties.pictureUrl = require("./assets/iconPerson.png");  
   }

   console.log("User's Email from pageContext: ", this.properties.UserEmail);
    console.log("User's Name from List: ", this.properties.userName);
    console.log("User's Profile from List: ", this.properties.pictureUrl); 
  }

  public getItemsFromSPList(listName: string): Promise<any[]> {
    return new Promise((resolve, reject) => {
      let open = indexedDB.open("MyDatabase", 1);
   
      open.onsuccess = function() {
        let db = open.result;
        let tx = db.transaction(`${listName}`, "readonly");
        let store = tx.objectStore(`${listName}`);
   
        let getAllRequest = store.getAll();
   
        getAllRequest.onsuccess = function() {
          resolve(getAllRequest.result);
        };
   
        getAllRequest.onerror = function() {
          reject(getAllRequest.error);
        };
      };
   
      open.onerror = function() {
        reject(open.error);
      };
    });
  }

  private async userDetails(): Promise<void> {
    // Ensure that you have access to the SPHttpClient
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
  
    // Use try-catch to handle errors
    try {
      // Get the current user's information
      const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
      const userProperties: any = await response.json();
  
      // console.log("User Details:", userProperties);
  
      // Access the userPrincipalName from userProperties
      const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
  
      if (userPrincipalNameProperty) {
        this.properties.userEmail = userPrincipalNameProperty.Value.toLowerCase();
        console.log('User Email using User Principal Name:', this.properties.userEmail);
      //   const pictureUrl: string = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?UserName=${this.properties.userEmail}`;
      //   console.log("User Profile Image from microsoft id:", pictureUrl);
      //   if(pictureUrl){
      //     this.properties.pictureUrl = pictureUrl;
      //   }else{
      //     this.properties.pictureUrl = require("./assets/iconPerson.png");
      //     console.error('PictureUrl property not found in user properties');
      //   }
      // const userNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'PreferredName');
      //   if(userNameProperty){
      //     this.properties.userName = userNameProperty.Value;
      //     console.log('User Name using User Principal Name:', this.properties.userName);
      //   }else{
      //     console.error('userPrincipalNameProperty property not found in user properties');
      //   }

      // console.log("this.properties.userEmail.includes('_') && this.properties.userEmail.includes('#ext#'", this.properties.userEmail.includes("_") && this.properties.userEmail.includes("#ext#"));
        if (this.properties.userEmail.includes("_") && this.properties.userEmail.includes("#ext#")){
          this.externalUser = true;
        }
      } else {
        console.error('User Principal Name not found in user properties');
      }
    } catch (error) {
      console.error('Error fetching user properties:', error);
    }
  } 
    
    private async getViVaEngageToken(){
      const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();    
      await tokenProvider.getToken("https://api.yammer.com").then(async token => {
        this.vivaEngageToken = token;
      });
    }

    private getGraphApiBookmarks(): Promise<any[]> {
      if (this.externalUser) {
        // If externalUser is true, return a resolved Promise with an empty array
        return Promise.resolve([]);
      }
    
      return new Promise<any[]>((resolve, reject) => {
        this.context.msGraphClientFactory
          .getClient('3')
          .then((client: MSGraphClientV3): void => {
            // get information about the current user from the Microsoft Graph
            client
              .api(`https://graph.microsoft.com/v1.0/drive/following?$filter=startsWith(webUrl,'${this.context.pageContext.web.absoluteUrl}')`)
              .get((error, rawResponse?: any) => {
                if (error) {
                  console.error(error);
                  reject(error);
                } else {
                  const bookmarks: any[] = rawResponse.value;
                  // console.log("Bookmarks Api Response:", bookmarks);
    
                  // Map the array to contain only 'name' and 'webUrl' properties
                  const simplifiedBookmarks = bookmarks.map(bookmark => ({
                    name: bookmark.name,
                    webUrl: bookmark.webUrl
                  }));
    
                  console.log("Simplified bookmarks response:", simplifiedBookmarks);
                  this.graphApiBookmarks = simplifiedBookmarks;
                  resolve(simplifiedBookmarks);
                }
              });
          })
          .catch((error) => {
            console.error(error);
            this.graphApiBookmarks = [];
            resolve([]); // Resolve with an empty array if an error occurs
          });
      });
    }

    private async userInfoWithSuccessFactorsList(): Promise<void> {
      // Get the current user's email
      let userEmail = this.context.pageContext.user.email.toLowerCase();
      console.log("User Email:", userEmail);
      if(!userEmail){
        let extracted = this.properties.userEmail.split("#ext#")[0];
        let lastUnderscoreIndex = extracted.lastIndexOf("_");
        if (lastUnderscoreIndex !== -1) {
          extracted = extracted.substring(0, lastUnderscoreIndex) + "@" + extracted.substring(lastUnderscoreIndex + 1);
        }
        userEmail = extracted;
        console.log(extracted);
      }
      // Construct the URL for the SharePoint REST API
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('SuccessFactors')/Items?$filter=Email eq '${userEmail}'&$select=UserName,ProfilePicture`;
      // console.log("Api Url:", apiUrl);
      // Fetch the items from the SharePoint list
      const response = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      const data = await response.json();
  
      // Log the items
      console.log("User from Sharepoint list:", data.value);
  
      if(data.value.length > 0){
        if(data.value[0].ProfilePicture){
      this.properties.pictureUrl = data.value[0].ProfilePicture.Url;
        }else{
          this.properties.pictureUrl = require("./assets/iconPerson.png");
        }
        if(data.value[0].UserName){
      this.properties.userName = data.value[0].UserName;
        }else{        
          this.properties.userName = "Welcome, User";
        }
      }else{
        this.properties.userName = "Welcome, User";
        this.properties.pictureUrl = require("./assets/iconPerson.png");
      }
    }
    
  private isRunningOnTeams = false;
  private isBodyEmbedded = false;

  public async onInit() {
    try {
      await app.initialize();
      const context = await app.getContext();
      console.log("Context:", context);
      if(context.app.host.name.includes("teams") || context.app.host.name.includes("Teams")){
        console.log("The extension is running inside Microsoft Teams");
        this.isRunningOnTeams = true;
      }else{
        console.log("The extension is running outside Microsoft Teams");
      }
    } catch (exp) {
        console.log("The extension is running outside Microsoft Teams");
    }
    this.isBodyEmbedded = document.body.classList.contains('embedded');
    if (this.isBodyEmbedded) {
      console.log('Body has the embedded class');
    } else {
      console.log('Body does not have the embedded class');
    }
      await this.userDetails();
      // this.users = await this.getItemsFromSPList("SPList");
      // await this.userInformation();
      await this.userInfoWithSuccessFactorsList();
      console.log('User details fetched successfully.');
      // await this.getViVaEngageToken();
      console.log("External User:", this.externalUser);
      if(!this.externalUser){
      await this.getGraphApiBookmarks();
      console.log('User bookmarks fetched successfully.');
      }

      this.context.placeholderProvider.changedEvent.add(this, this.renderNavbar);
  }

  public renderNavbar(): void{
    // console.log("Start of render");
    // console.log(
    //   "Available placeholders: ",
    //   this.context.placeholderProvider.placeholderNames
    //     .map(name => PlaceholderName[name])
    //     .join(", ")
    // );
    const profilePhotoUrl = this.properties.pictureUrl.length > 0 ? this.properties.pictureUrl : `${require("./assets/iconPerson.png")}`;
    const userName = this.properties.userName.length > 0 ? this.properties.userName: '';
    
    const dropdownIcon = require("./assets/dropdownIcon.png");
    const navIcon = require("./assets/nav-icon.png");
    const navIconClose = require("./assets/nav-icon-close.png");
    const facebookIcon = require("./assets/facebook.png");
    const instagramIcon = require("./assets/instagram.png");
    const linkdinIcon = require("./assets/linkdin.png");
    const twitterIcon = require("./assets/twitter.png");
    const youtubeIcon = require("./assets/youtube.png");
    const rpgLogo = require("./assets/rpg-logo.png");
    let homeHref = "";
    let rpgSocialHref = "";

    if(this.isRunningOnTeams && this.isBodyEmbedded){
      console.log("The site is running on teams");
      homeHref = "https://rpgnet.sharepoint.com/sites/OneRPG/SitePages/HomeTeams.aspx";
      rpgSocialHref = "https://rpgnet.sharepoint.com/sites/OneRPG/SitePages/RPG-SocialTeams.aspx";
    }else if(!this.isRunningOnTeams && this.isBodyEmbedded){
      console.log("The site is neither running on teams nor sharepoint site");
      homeHref = "https://rpgnet.sharepoint.com/sites/OneRPG";
      rpgSocialHref = "https://rpgnet.sharepoint.com/sites/OneRPG/SitePages/RPG-Social.aspx";
    }else{
      console.log("The site is running on the sharepoint site")
    }

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this.onDispose.bind(this) } // Ensure onDispose is bound to 'this'
      );

  
      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }
  
      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = "(Top property was not defined.)";
        }
  
        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
    
    <div class="${styles.navRightSection} navRightSection">
    <div class="${styles.contents} navContents">
    <div class="${styles.navIcons}">
    <a class="${styles.HomeBtn} homeBtn" href="${homeHref}">Home</a>
    <a class="${styles.RpgSocialBtn} homeBtn" href="${rpgSocialHref}">RPG Social</a>
    <a class="${styles.oneRPG}" href="${this.context.pageContext.web.absoluteUrl}/SitePages/Know-OneRPG.aspx" target="_self" data-interception="off">Know OneRPG</a>
    
    </div>

    <div class="${styles.navIcons}">

            <div class="${styles.dropDownIcon}">
               <a class="${styles.navDropIcon}" id="myApplicationLink">My Applications <img src="${dropdownIcon}"></a> 
               <ul class="${styles.navDropIconBox}  dropBoxMyapp">
               </ul>
            </div>

            <div class="${styles.dropDownIcon}" id="Bookmark">
               <a class="${styles.navDropIcon}" id="MyBookmarksLink">My Bookmarks  <img src="${dropdownIcon}"></a> 
               <ul class="${styles.navDropIconBox} dropBoxMyBook">
               </ul>
            </div>
            
        </div>

        <div class="${styles.notification}">
            <a class="${styles.notificationIcon} notificationImage" id="notificationLink" >
            </a>
                <ul class="${styles.navDropIconBox} notificationIconBox">
               </ul>
        </div>

        <div class="${styles.profile}">
          <a id="userName">${userName}</a>
            <div class="${styles.imgBox}">
                <a href="${this.context.pageContext.web.absoluteUrl}/SitePages/People-Connect.aspx">
                 <img src="${profilePhotoUrl}">
                </a>
            </div>
        </div>
    </div>

    <div class="${styles.iconMenuNav}">
        <i class="${styles.navIcon}" id="navIcon"><img src="${navIcon}"></i>
        <i class="${styles.closeBtn}" id="closeBtn"><img src="${navIconClose}"> </i>
    </div>

</div>

    `;

    const link = document.querySelector("link[rel~='icon']") as HTMLLinkElement;

  if (!link) {
    const newLink = document.createElement('link') as HTMLLinkElement;
    newLink.rel = 'icon';
    document.getElementsByTagName('head')[0].appendChild(newLink);
  }

  // Set the href attribute to the URL of your favicon
  link.href = 'https://rpgnet.sharepoint.com/sites/OneRPG/SiteAssets/favicon.png';



        }
      }
    }
  
   if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }
  
      if (this.properties) {
        let bottomString: string = this.properties.Bottom;
        if (!bottomString) {
          bottomString = "(Bottom property was not defined.)";
        }
  
        if (this._bottomPlaceholder.domElement) {
          var targetElement = $('footer > div >  div:first-child ');
 
        // Check if the target element exists
        if (targetElement.length > 0) {
            // Add HTML content to the first child of the second level of div
            targetElement.html('<div class="footerContents"> <img src="https://rpgnet.sharepoint.com/sites/OneRPG_Test/SiteAssets/images/rpg-logo.png" class="footer_logo"> <div class="social_icons"> <a href="https://twitter.com/rpgenterprises" target="blank"><img src="https://rpgnet.sharepoint.com/sites/OneRPG_Test/SiteAssets/images/social-media-icons/twitter.png" class="twitter_icon"></a> <a href="https://www.facebook.com/RPGGroup" target="blank"><img src="https://rpgnet.sharepoint.com/sites/OneRPG_Test/SiteAssets/images/social-media-icons/facebook.png" class="facebook_icon"></a> <a href="https://www.instagram.com/rpg.group/" target="blank"><img src="https://rpgnet.sharepoint.com/sites/OneRPG_Test/SiteAssets/images/social-media-icons/instagram.png" class="insta_icon"></a> <a href="#" target="blank"><img src="https://rpgnet.sharepoint.com/sites/OneRPG_Test/SiteAssets/images/social-media-icons/youtube.png" class="youtube_icon"></a> <a href="https://www.linkedin.com/company/rpg-group/" target="blank"><img src="https://rpgnet.sharepoint.com/sites/OneRPG_Test/SiteAssets/images/social-media-icons/linkdin.png" class="linkdin_icon"></a> </div> </div>');
        }

        targetElement.find('.social_icons').addClass(styles.socialOIcons);
        targetElement.find('.footer_logo').addClass(styles.footerLogo);
        targetElement.find('.footerContents').addClass(styles.footerContents);
        targetElement.find('.youtube_icon').addClass(styles.youtubeIcon);
        targetElement.find('.linkdin_icon').addClass(styles.linkdinIcon);
        
        }
      }
    }  
  this.navBarItems();
  }

  private executedPreviousMethods: boolean = false;

  private async navBarItems(): Promise<void> {
    if (!this.executedPreviousMethods) {
      await this.renderButtonsApplication();
      await this.renderButtonsBookmarks(this.graphApiBookmarks); 
      await this.renderButtonsNotifications();    
      // await this.getGraphApiNotifications();
      // await this.getYammerNotifications();
  
      this.executedPreviousMethods = true; // Set the flag to indicate that previous methods have been executed
    }
  
    // this.checkIfEmpty();
    this.addEventHandlers();
  }

  // private async getGraphApiNotifications(): Promise<void> {

  //   const today = new Date();
  //   const TwoDaysEarlier = new Date(today);
  //   TwoDaysEarlier.setDate(today.getDate() - 2);
  //   // console.log("today: ", today);
  //   // console.log("TwoDaysEarlier: ", TwoDaysEarlier);
  //   const formattedToday = `${today.getFullYear().toString()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
  //   const formattedTwoDaysEarlier = `${TwoDaysEarlier.getFullYear().toString()}-${(TwoDaysEarlier.getMonth() + 1).toString().padStart(2, '0')}-${TwoDaysEarlier.getDate().toString().padStart(2, '0')}`;
  //   // console.log("formattedToday: ", formattedToday);
  //   // console.log("formattedTwoDaysEarlier: ", formattedTwoDaysEarlier);

  //   const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20&Modified%20ge%20%27${formattedTwoDaysEarlier}%27&$select=ID,FileLeafRef,Title`;
  //   await fetch(apiUrl, {
  //     method: 'GET',
  //     headers: {
  //       'Accept': 'application/json;odata=nometadata',
  //       'Content-Type': 'application/json;odata=nometadata',
  //       'odata-version': ''
  //     }
  //   })
  //   .then(response => response.json())
  //   .then(data => {
  //     console.log("Notifications from sharepoint response: ", data);
      
  //     if (data.value && data.value.length > 0) {
  //       data.value.forEach(async (ArticleItem: SharepointNotifications) => {
  //         const apiCall: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/GetItemById(${ArticleItem.ID})/Comments`;
  //         await fetch(apiCall, {
  //           method: 'GET',
  //           headers: {
  //             'Accept': 'application/json;odata=nometadata',
  //             'Content-Type': 'application/json;odata=nometadata',
  //             'odata-version': ''
  //           }
  //         })
  //         .then(response => response.json())
  //         .then(data => {
  //           // console.log(`Comments found for id ${ArticleItem.ID}:`, data);
  //           if (data.value && data.value.length > 0) {
  //             data.value.forEach(async (CommentItem: SharepointNotifications) => {
  //               // console.log("CommentItem: ", CommentItem);
  //               if (CommentItem.text.includes(this.properties.userEmail)){
  
  //                 // Filter items based on month and date on the client side
  //                 if (
  //                   CommentItem.createdDate &&
  //                   CommentItem.createdDate.substring(0, 10) >= formattedTwoDaysEarlier &&
  //                   CommentItem.createdDate.substring(0, 10) <= formattedToday
  //               ) {
  //                   // console.log('Creating button for ', ArticleItem.Title);
  //                   const button: HTMLLIElement = document.createElement('li');
  //                   button.classList.add(styles.noMargin);
  //                   button.textContent = 'You have been mentioned in ' + ArticleItem.Title;
  //                   button.onclick = () => {
  //                       window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${ArticleItem.FileLeafRef}`, '_self');
  //                   };
  //                   const notificationIconBox = document.querySelector('.notificationIconBox');
  //                   if (notificationIconBox) {
  //                       notificationIconBox.appendChild(button);
  //                       this.noOfNotifications++;
  //                   } else {
  //                       console.warn('Element with class "notificationIconBox" not found.');
  //                   }
  //               } else {
  //                   // console.log('There are no mentions for this user the last 3 days');
  //               }
                  
  //                 } else {
  //                 //  console.log("There are no mentions for this user");
  //                 }
  //             });
             
  //             } else {
  //               // console.log("There are no mentions for this article");
  //             }
  //         })
  //         .catch(error => {
  //           console.error("Error fetching user data: ", error);
  //         });
  //       });
  //     } else {
  //       // console.log("There are no mentions for this user");
  //     }
  //   })
  //   .catch(error => {
  //     console.error("Error fetching user data: ", error);
  //   });
  // }

  private setNotifColor = true;

  private async renderButtonsApplication(): Promise<void> {
    // console.log("User's Email from LoginName: ", this.properties.userEmail);
    const adminEmailSplit: string[] = this.properties.userEmail.split('.admin@');
    if (this.properties.userEmail.includes(".admin@")){
      console.log("Admin Email after split: ", adminEmailSplit);
    }    
    let otherUsersSplit = "";
    if (this.properties.userEmail.includes("_") && this.properties.userEmail.includes("#ext#")){
    const parts = this.properties.userEmail.split('_');
    const secondPart = parts.length > 1 ? parts[1] : '';
    otherUsersSplit =  secondPart.split('.com')[0];
      console.log("User's company after split: ", otherUsersSplit);
    }

    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Application_List')/items`;

    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
    .then(response => response.json())
    .then(data => {
      console.log("Api response: ", data);
      let buttonsCreated = 0; // Variable to keep track of the number of buttons created

      if (data.value && data.value.length > 0) {
        data.value.forEach((item: ApplicationsListItem) => {

          if((this.properties.userEmail.includes("@"+item.Company.toLowerCase()+".") && !this.properties.userEmail.includes(".admin@") && !otherUsersSplit) || (this.properties.userEmail.includes(".admin@") && adminEmailSplit.includes("@"+item.Company.toLowerCase()+".")) || (otherUsersSplit.length >= 0 && otherUsersSplit.includes(item.Company.toLowerCase()))){                    
            // console.log("Creating button for ", item.Title);
           const button: HTMLLIElement = document.createElement('li');
            button.classList.add(styles.noMargin); 
            button.textContent = item.Title; 
            button.onclick = () => {
              window.open(item.URL, '_blank'); 
            };
            const dropBoxMyappElement = document.querySelector('.dropBoxMyapp');

            if (dropBoxMyappElement) {
              dropBoxMyappElement.appendChild(button);
              this.noOfApplications++
            } else {
              console.warn('Element with class "dropBoxMyapp" not found.');
            }
            buttonsCreated++; 
          } else {
            // console.log("No applications available for the user");
          }
        });
      } else {
        // console.log("No applications available for the user");
      }
      
      if(this.noOfApplications <= 0){
        console.log("No applications available for the user");
        const dropBoxMyappElement = document.querySelector('.dropBoxMyapp');
        dropBoxMyappElement.innerHTML = '';
        const noDataMessage: HTMLLIElement = document.createElement('li');            
        noDataMessage.classList.add(styles.noMargin); 
        noDataMessage.textContent = 'No applications available for the user.';
  
        if (dropBoxMyappElement) {
          dropBoxMyappElement.appendChild(noDataMessage);
          console.log("Displaying no Applications");
        } else {
          console.warn('Element with class "dropBoxMyapp" not found.');
        }
      }
    })
    .catch(error => {
      console.error("Error fetching user data: ", error);
    });
  }

  private async renderButtonsBookmarks(bookmarks: any[]): Promise <void> {
    const dropBoxMyBookElement = document.querySelector('.dropBoxMyBook') as HTMLElement;
    const bookmarkTab = document.getElementById('Bookmark') as HTMLDivElement;

    // here edit for visibility of the bookmark tab of external users
    // CURRENTLY NOT WORKING
    if (this.externalUser) {
      console.log("No bookmarks available for the user");
      if (bookmarkTab) {
        console.warn('Element with id "Bookmark" found.');
        bookmarkTab.style.display = 'none';
      } else {
        console.warn('Element with id "Bookmark" not found.');
      }
      return;
    }
    //end of edit

    if (bookmarks && bookmarks.length > 0){
    if (dropBoxMyBookElement) {
      bookmarks.forEach((bookmark: any) => {
        const button: HTMLLIElement = document.createElement('li');
        button.classList.add(styles.noMargin); 

        button.textContent = bookmark.name;
        button.onclick = () => {
          window.open(bookmark.webUrl, '_self');
        };
        // console.log("Added button for:", bookmark.name);
        dropBoxMyBookElement.appendChild(button);
        this.noOfBookmarks++;
      });
    } else {
      // console.warn('Element with class "dropBoxMyBook" not found.');
    }
    }else{
      // console.log("No bookmarks available for the user");
    }
    if(this.noOfBookmarks === 0){
      console.log("No bookmarks available for the user");
      const noDataMessage: HTMLLIElement = document.createElement('li');            
        noDataMessage.classList.add(styles.noMargin); 
        noDataMessage.textContent = 'No bookmarks available for the user.';
        const dropBoxMyBookElement = document.querySelector('.dropBoxMyBook')
        dropBoxMyBookElement.innerHTML = '';
        if (dropBoxMyBookElement) {
          dropBoxMyBookElement.appendChild(noDataMessage);
          console.log("Displaying no Bookmarks");
        } else {
        console.warn('Element with class "dropBoxMyBookElement" not found.');
        }
      }
  }

  private async getYammerNotifications() {
    const apiUrl: string = 'https://api.yammer.com/api/v1/messages/received.json';
  
    try {
      const response = await fetch(apiUrl, {
        method: 'GET',
        headers: {
          "Authorization": `Bearer ${this.vivaEngageToken}`,
          'Content-Type': 'application/json;odata=nometadata',
        }
      });
  
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
  
      const data = await response.json();
      console.log("Yammer Notifications response:", data);
  
      const today = new Date();
      const twoDaysAgo = new Date(today);
      twoDaysAgo.setDate(today.getDate() - 2);
  
      if (data.messages && data.messages.length > 0) {
        data.messages.forEach((message: YammerNotifications) => {
          if (message.body.plain) {
            const messageDate = new Date(message.published_at.substring(0,10));
            
            if (messageDate >= twoDaysAgo && messageDate <= today) {
              console.log("Creating button for Yammer message:", message.body.plain);
  
              const button: HTMLLIElement = document.createElement('li');
              button.classList.add(styles.noMargin);
              button.textContent = "You have a new Engage mention: " + message.body.plain;
              button.onclick = () => {
                window.open(message.web_url, '_blank');
              };
  
              const notificationIconBox = document.querySelector('.notificationIconBox');
              if (notificationIconBox) {
                notificationIconBox.appendChild(button);
                this.noOfNotifications++;
              } else {
                console.warn('Element with class "notificationIconBox" not found.');
              }
            }
          }
        });
      } else {
        console.log("There are no new Yammer mentions");
      }
    } catch (error) {
      console.error("Error fetching Yammer notifications:", error);
    }
  }
  
  private async renderButtonsNotifications(): Promise<void> {
    
    try{

    const adminEmailSplit: string[] = this.properties.userEmail.split('.admin@');
    if (this.properties.userEmail.includes(".admin@")){
        console.log("Admin Email after split: ", adminEmailSplit);
    }
   let otherUsersSplit = "";
    if (this.properties.userEmail.includes("_") && this.properties.userEmail.includes("#ext#")){
    const parts = this.properties.userEmail.split('_');
    const secondPart = parts.length > 1 ? parts[1] : '';
    otherUsersSplit =  secondPart.split('.com')[0];
      console.log("User's company after split: ", otherUsersSplit);
    }

    const today = new Date();
    const TwoDaysEarlier = new Date(today);
    TwoDaysEarlier.setDate(today.getDate() - 2);
    // console.log("today: ", today);
    // console.log("TwoDaysEarlier: ", TwoDaysEarlier);
    const formattedToday = `${today.getFullYear().toString()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
    const formattedTwoDaysEarlier = `${TwoDaysEarlier.getFullYear().toString()}-${(TwoDaysEarlier.getMonth() + 1).toString().padStart(2, '0')}-${TwoDaysEarlier.getDate().toString().padStart(2, '0')}`;
    // console.log("formattedToday: ", formattedToday);
    // console.log("formattedTwoDaysEarlier: ", formattedTwoDaysEarlier);

  let apiUrl: string = ``;
if (this.properties.userEmail.includes(".admin@")) {
    apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Notification_Panel')/items?$filter=(TargetAudience eq '16' or TargetAudience eq '${this._checkIfAdmin()}' or TargetAudience eq '${this._getCompanyFromEmail()}') and (StartDate le '${formattedToday}' and StartDate ge '${formattedTwoDaysEarlier}') &$orderby=StartDate`;
} else {
    apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Notification_Panel')/items?$filter=(TargetAudience eq '16' or TargetAudience eq '${this._getCompanyFromEmail()}') and (StartDate le '${formattedToday}' and StartDate ge '${formattedTwoDaysEarlier}') &$orderby=StartDate`;
}


    fetch(apiUrl, {
        method: 'GET',
        headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': ''
        }
    })
    .then(response => response.json())
    .then(data => {
        console.log("Api response: ", data);
        let buttonsCreated = 0; // Variable to keep track of the number of buttons created
        if (data.value && data.value.length > 0) {
            data.value.forEach((item: NotificationsListItem) => {

              if((this.properties.userEmail.includes("@"+item.Company.toLowerCase()+".") && !this.properties.userEmail.includes(".admin@") && !otherUsersSplit) || (this.properties.userEmail.includes(".admin@") && adminEmailSplit.includes("@"+item.Company.toLowerCase()+".")) || (otherUsersSplit.length >= 0 && otherUsersSplit.includes(item.Company.toLowerCase()))){                    
                // console.log("Creating button for ", item.NotificationText);
                      const button: HTMLLIElement = document.createElement('li');
                      button.classList.add(styles.noMargin); 
                      button.textContent = item.NotificationText;
                      button.onclick = () => {
                        window.open(item.NotificationURL.Url, '_blank'); // Open the 'Url' from the API response in a new tab
                    };
                    const notificationIconBox = document.querySelector('.notificationIconBox');
                    if (notificationIconBox) {
                      notificationIconBox.appendChild(button);
                    this.noOfNotifications++;
                    } else {
                      console.warn('Element with class "notificationIconBox" not found.');
                    }buttonsCreated++; // Increment the count of buttons created
                } else {
                    // console.log("No button creation for: ", item.NotificationText);
                }
            });
        } else {
          console.log("There are no Notifications from admin for user");
        }
    })
    .catch(error => {
        console.error("Error fetching user data: ", error);
    });

    const ApiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/items?$filter=PromotedState%20eq%202%20&Modified%20ge%20%27${formattedTwoDaysEarlier}%27&$select=ID,FileLeafRef,Title`;
    await fetch(ApiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
    .then(response => response.json())
    .then(data => {
      console.log("Notifications from sharepoint response: ", data);
      
      if (data.value && data.value.length > 0) {
        data.value.forEach(async (ArticleItem: SharepointNotifications) => {
          const apiCall: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site%20Pages')/GetItemById(${ArticleItem.ID})/Comments`;
          await fetch(apiCall, {
            method: 'GET',
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          })
          .then(response => response.json())
          .then(data => {
            // console.log(`Comments found for id ${ArticleItem.ID}:`, data);
            if (data.value && data.value.length > 0) {
              data.value.forEach(async (CommentItem: SharepointNotifications) => {
                // console.log("CommentItem: ", CommentItem);
                if (CommentItem.text.includes(this.properties.userEmail)){
  
                  // Filter items based on month and date on the client side
                  if (
                    CommentItem.createdDate &&
                    CommentItem.createdDate.substring(0, 10) >= formattedTwoDaysEarlier &&
                    CommentItem.createdDate.substring(0, 10) <= formattedToday
                ) {
                    // console.log('Creating button for ', ArticleItem.Title);
                    const button: HTMLLIElement = document.createElement('li');
                    button.classList.add(styles.noMargin);
                    button.textContent = 'You have been mentioned in ' + ArticleItem.Title;
                    button.onclick = () => {
                        window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/${ArticleItem.FileLeafRef}`, '_self');
                    };
                    const notificationIconBox = document.querySelector('.notificationIconBox');
                    if (notificationIconBox) {
                        notificationIconBox.appendChild(button);
                        this.noOfNotifications++;
                    } else {
                        console.warn('Element with class "notificationIconBox" not found.');
                    }
                } else {
                    // console.log('There are no mentions for this user the last 3 days');
                }
                  
                  } else {
                  //  console.log("There are no mentions for this user");
                  }
              });
             
              } else {
                // console.log("There are no mentions for this article");
              }
          })
          .catch(error => {
            console.error("Error fetching user data: ", error);
          });
        });
      } else {
        // console.log("There are no mentions for this user");
      }
    })
    .catch(error => {
      console.error("Error fetching user data: ", error);
    });

  let notifications = true;

  if(this.noOfNotifications === 0){
    notifications=false;
    console.log("There a no Notifications");
    const noDataMessage: HTMLLIElement = document.createElement('li');            
    noDataMessage.classList.add(styles.noMargin); 
    noDataMessage.textContent = 'No notifications available for the user';
    const notificationIconBox = document.querySelector('.notificationIconBox');
    notificationIconBox.innerHTML = '';
    if (notificationIconBox) {
      notificationIconBox.appendChild(noDataMessage);
      console.log("Displaying no Notifications");
    } else {
      console.warn('Element with class "notificationIconBox" not found.');
    }
  }

  let setNotifIcon = true;
  if(setNotifIcon){
    // console.log("Setting Notification Icon"); 
  const NotificationIcon = require("./assets/notification-icon.png");
  const noNotificationIcon = require("./assets/no-notification-icon.png");

  const notificationIcon = document.querySelector('.notificationImage');
  notificationIcon.innerHTML = '';
  const notifIcon: HTMLImageElement = document.createElement('img');

//  console.log("Notif found", notificationIcon);
  if (notifications) {
    // console.log("Notification Icon displaying");
    notifIcon.src = NotificationIcon;
  } else {
    // console.log("NoNotification Icon displaying");      
    notifIcon.src = noNotificationIcon;
    this.setNotifColor = false;
  }
  notificationIcon.appendChild(notifIcon);
  setNotifIcon = false;
  }
  
  }catch (error) {
    console.error('Error fetching user properties:', error);
  }

}

private async checkIfEmpty(): Promise<void> {
  let notifications = true;

  // if(this.noOfBookmarks === 0){
  //   console.log("There a no Bookmarks");
  //   const noDataMessage: HTMLLIElement = document.createElement('li');            
  //   noDataMessage.classList.add(styles.noMargin); 
  //   noDataMessage.textContent = 'No bookmarks available for the user.';
  //   const dropBoxMyBookElement = document.querySelector('.dropBoxMyBook')
  //   dropBoxMyBookElement.innerHTML = '';
  //   if (dropBoxMyBookElement) {
  //     dropBoxMyBookElement.appendChild(noDataMessage);
  //     console.log("Displaying no Bookmarks");
  //   } else {
  //   console.warn('Element with class "dropBoxMyBookElement" not found.');
  //   }
  // }
  // if(this.noOfApplications === 0){
  //   console.log("There a no Applications");
  //   const dropBoxMyappElement = document.querySelector('.dropBoxMyapp');
  //   dropBoxMyappElement.innerHTML = '';
  //   const noDataMessage: HTMLLIElement = document.createElement('li');            
  //   noDataMessage.classList.add(styles.noMargin); 
  //   noDataMessage.textContent = 'No applications available for the user.';

  //   if (dropBoxMyappElement) {
  //     dropBoxMyappElement.appendChild(noDataMessage);
  //     console.log("Displaying no Applications");
  //   } else {
  //     console.warn('Element with class "dropBoxMyapp" not found.');
  //   }
  // }

//   if(this.noOfNotifications === 0){
//     notifications=false;
//     console.log("There a no Notifications");
//     const noDataMessage: HTMLLIElement = document.createElement('li');            
//     noDataMessage.classList.add(styles.noMargin); 
//     noDataMessage.textContent = 'No notifications available for the user';
//     const notificationIconBox = document.querySelector('.notificationIconBox');
//     notificationIconBox.innerHTML = '';
//     if (notificationIconBox) {
//       notificationIconBox.appendChild(noDataMessage);
//       console.log("Displaying no Notifications");
//     } else {
//       console.warn('Element with class "notificationIconBox" not found.');
//     }
//   }

//   let setNotifIcon = true;
//   if(setNotifIcon){
//     // console.log("Setting Notification Icon"); 
//   const NotificationIcon = require("./assets/notification-icon.png");
//   const noNotificationIcon = require("./assets/no-notification-icon.png");

//   const notificationIcon = document.querySelector('.notificationImage');
//   notificationIcon.innerHTML = '';
//   const notifIcon: HTMLImageElement = document.createElement('img');

// //  console.log("Notif found", notificationIcon);
//   if (notifications) {
//     // console.log("Notification Icon displaying");
//     notifIcon.src = NotificationIcon;
//   } else {
//     // console.log("NoNotification Icon displaying");      
//     notifIcon.src = noNotificationIcon;
//     this.setNotifColor = false;
//   }
//   notificationIcon.appendChild(notifIcon);
//   setNotifIcon = false;
//   }
}
  
  private addEventHandlers(): void {
    // Add event handler here
    const myApplicationLink = this._topPlaceholder.domElement.querySelector('#myApplicationLink');
    const MyBookmarksLink = this._topPlaceholder.domElement.querySelector('#MyBookmarksLink');
    const notificationLink = this._topPlaceholder.domElement.querySelector('#notificationLink');
    const navIcon = this._topPlaceholder.domElement.querySelector('#navIcon');
    const closeBtn = this._topPlaceholder.domElement.querySelector('#closeBtn');
  
    if (myApplicationLink) {
      myApplicationLink.addEventListener('click', this.handleMyApplicationClick);
    }
    if (MyBookmarksLink) {
      MyBookmarksLink.addEventListener('click', this.handleMyBookmarksClick);
    }
    if (notificationLink) {
      notificationLink.addEventListener('click', this.handlenotificationLink);
    }
    if (navIcon) {
      navIcon.addEventListener('click', () => this.handlenavIconClick()); // Wrap the function call
    }
    if (closeBtn) {
      closeBtn.addEventListener('click', this.handlecloseBtnClick);
    }
  
    // Add a click event listener to the body to remove classes on body click
    document.body.addEventListener('click', this.handleBodyClick);
  } 

  private handleBodyClick = (event: MouseEvent): void => {
  const myApplicationLink = this._topPlaceholder.domElement.querySelector('#myApplicationLink');
  const MyBookmarksLink = this._topPlaceholder.domElement.querySelector('#MyBookmarksLink');
  const dropBoxMyapp = this._topPlaceholder.domElement.querySelector('.dropBoxMyapp');
  const dropBoxMyBook = this._topPlaceholder.domElement.querySelector('.dropBoxMyBook');
  const notificationLink = this._topPlaceholder.domElement.querySelector('#notificationLink');
  const notificationIconBox = this._topPlaceholder.domElement.querySelector('.notificationIconBox');

  // Check if the click target is outside of myApplicationLink and dropBoxMyapp
  if (
    myApplicationLink &&
    MyBookmarksLink &&
    notificationLink &&
    dropBoxMyapp &&
    dropBoxMyBook &&
    notificationIconBox &&
    !myApplicationLink.contains(event.target as Node) &&
    !MyBookmarksLink.contains(event.target as Node) &&
    !notificationLink.contains(event.target as Node) &&
    !dropBoxMyapp.contains(event.target as Node) &&
    !dropBoxMyBook.contains(event.target as Node) &&
    !notificationIconBox.contains(event.target as Node)
  ) {
    myApplicationLink.classList.remove(styles.activeNav);
    MyBookmarksLink.classList.remove(styles.activeNav);
    notificationLink.classList.remove(styles.activeNav);
    dropBoxMyapp.classList.remove(styles.dBlock);
    dropBoxMyBook.classList.remove(styles.dBlock);
    notificationIconBox.classList.remove(styles.dBlock);
  }
};

private handlenavIconClick = (): void => {
  // console.log('navIconClick handler called');
  const navIcon = this._topPlaceholder.domElement.querySelector('#navIcon');
  const navContents = this._topPlaceholder.domElement.querySelector(".navContents");
  const cancelBtn = this._topPlaceholder.domElement.querySelector("#closeBtn");

    navContents.classList.add(styles.navShow);
    navIcon.classList.add(styles.dNone);
    cancelBtn.classList.add(styles.dBlock);
  
};


private handlecloseBtnClick = (): void => {
  const navIcon = this._topPlaceholder.domElement.querySelector('#navIcon');
  const navContents = this._topPlaceholder.domElement.querySelector(".navContents");
  const cancelBtn = this._topPlaceholder.domElement.querySelector("#closeBtn");

    navContents.classList.remove(styles.navShow);
    navIcon.classList.remove(styles.dNone);
    cancelBtn.classList.remove(styles.dBlock);

};

private handleMyApplicationClick = (): void => {

  const myApplicationLink = this._topPlaceholder.domElement.querySelector('#myApplicationLink');
  const dropBoxMyapp = this._topPlaceholder.domElement.querySelector('.dropBoxMyapp');
  const MyBookmarksLink = this._topPlaceholder.domElement.querySelector('#MyBookmarksLink');
  const dropBoxMyBook = this._topPlaceholder.domElement.querySelector('.dropBoxMyBook');
  const notificationLink = this._topPlaceholder.domElement.querySelector('#notificationLink');
  const notificationIconBox = this._topPlaceholder.domElement.querySelector('.notificationIconBox');


  if (myApplicationLink && dropBoxMyapp && MyBookmarksLink && dropBoxMyBook && notificationLink && notificationIconBox) {
      // Check if the class is added
    myApplicationLink.classList.add(styles.activeNav);
    dropBoxMyapp.classList.toggle(styles.dBlock);
    MyBookmarksLink.classList.remove(styles.activeNav);
    dropBoxMyBook.classList.remove(styles.dBlock);
    notificationLink.classList.remove(styles.activeNav);
    notificationIconBox.classList.remove(styles.dBlock);
  } else {
    console.warn('Element not found');
  }
};

private handleMyBookmarksClick = (): void => {

  const MyBookmarksLink = this._topPlaceholder.domElement.querySelector('#MyBookmarksLink');
  const dropBoxMyBook = this._topPlaceholder.domElement.querySelector('.dropBoxMyBook');
  const myApplicationLink  = this._topPlaceholder.domElement.querySelector('#myApplicationLink');
  const dropBoxMyapp = this._topPlaceholder.domElement.querySelector('.dropBoxMyapp');
  const notificationLink = this._topPlaceholder.domElement.querySelector('#notificationLink');
  const notificationIconBox = this._topPlaceholder.domElement.querySelector('.notificationIconBox');


  if (MyBookmarksLink && dropBoxMyBook && myApplicationLink && dropBoxMyapp && notificationLink && notificationIconBox) {
      // Check if the class is added
      MyBookmarksLink.classList.add(styles.activeNav);
      dropBoxMyBook.classList.toggle(styles.dBlock);
      myApplicationLink.classList.remove(styles.activeNav);
      dropBoxMyapp.classList.remove(styles.dBlock);
      notificationLink.classList.remove(styles.activeNav);
      notificationIconBox.classList.remove(styles.dBlock);
  } else {
    console.warn('Element not found');
  }
};

private handlenotificationLink = (): void => {

  const MyBookmarksLink = this._topPlaceholder.domElement.querySelector('#MyBookmarksLink');
  const dropBoxMyBook = this._topPlaceholder.domElement.querySelector('.dropBoxMyBook');
  const myApplicationLink  = this._topPlaceholder.domElement.querySelector('#myApplicationLink');
  const dropBoxMyapp = this._topPlaceholder.domElement.querySelector('.dropBoxMyapp');
  const notificationLink = this._topPlaceholder.domElement.querySelector('#notificationLink');
  const notificationIconBox = this._topPlaceholder.domElement.querySelector('.notificationIconBox');


  if (MyBookmarksLink && dropBoxMyBook && myApplicationLink && dropBoxMyapp && notificationLink && notificationIconBox) {
     
    if(this.setNotifColor){
      notificationLink.classList.add(styles.activeNav);
    }
      notificationIconBox.classList.toggle(styles.dBlock);
      myApplicationLink.classList.remove(styles.activeNav);
      dropBoxMyapp.classList.remove(styles.dBlock);
      MyBookmarksLink.classList.remove(styles.activeNav);
      dropBoxMyBook.classList.remove(styles.dBlock); 
  } else {
    console.warn('Element not found');
  }    
}

private _checkIfAdmin(): string{
  let adminGroup: string = "";

  if(this.properties.userEmail.includes(".admin@")){
    console.log("This user is an admin");
    adminGroup = "19";
  }else{
    console.log("User is not Admin");
  }
  return adminGroup;

}

private _getCompanyFromEmail(): string {
  let userGroup: string = "";

  if(this.properties.userEmail.includes("@aciesinnovations.com"))
  {
    console.log("User belongs to acies");
    userGroup = "7";
  }else if(this.properties.userEmail.includes("_zensar.com") || this.properties.userEmail.includes("@zensar.")){
    console.log("User belongs to zensar");
    userGroup = "20";
  }else if(this.properties.userEmail.includes("@rpg.com") || this.properties.userEmail.includes("@rpg.in")){
    console.log("User belongs to rpg");
    userGroup = "17";
  }else if(this.properties.userEmail.includes("@ceat.com")){
    console.log("User belongs to ceat");
    userGroup = "36";
  }else if(this.properties.userEmail.includes("_harrisonsmalayalam.com") || this.properties.userEmail.includes("@harrisonsmalayalam.com")){
    console.log("User belongs to harrison");
    userGroup = "18";
  }else if(this.properties.userEmail.includes("@kecrpg.com")){
    console.log("User belongs to kec");
    userGroup = "39";
  }else if(this.properties.userEmail.includes("@raychemrpg.com")){
    console.log("User belongs to raychem");
    userGroup = "37";
  }else if(this.properties.userEmail.includes("@rpgls.com")){
    console.log("User belongs to rpgls");
    userGroup = "38";
  }
return userGroup;
}  

  private _onDispose(): void {
    console.log('[NavbarExtensionApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
  
}