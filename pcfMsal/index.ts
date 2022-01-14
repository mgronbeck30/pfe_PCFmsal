import {IInputs, IOutputs} from "./generated/ManifestTypes";
import {Authority, AuthResponse, Configuration, UserAgentApplication, }  from 'msal';
import * as msal2 from "@azure/msal-common";
import * as msal3 from "@azure/msal-node";
import * as ReactDOM from "react-dom";
import * as React from "react";
import * as button from './reactButton';
import * as $ from 'jquery';
import { networkInterfaces } from "os";
import { DeviceCodeResponse, DeviceCodeClient, CommonDeviceCodeRequest,CacheManager } from "@azure/msal-common";


//since this can be called from various powerapps pages (make.powerapps.com, us.create.powerapps.com, orgxxxx.crm.dynamics.com, etc) 
//we need a wide range of app registration redirect uri's
const redirect = window.location.host;
//const webAPIURL = "https://professionalservices.microsoft.com/lmt-coreapi/api/v1/assignments/danorsen?sourceSystemIds=bravos,bravos-servicecenter&IsActive=true"
const webAPIURL = "https://gronbeckd365.crm.dynamics.com/api/data/v9.1/new_sessionaudits"
/**
 * clientId will need to be changed to match the app registration you create in Azure AD
 * https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps 
 * authority needs to have your tenant id at the end, instead of mine
 * any possible value in redirect URI also needs to be registered with the app registratin, here are mine:
 * *https://apps.preview.powerapps.com
 * *https://gronbeckinc.crm.dynamics.com
 * *https://us.create.powerapps.com
 * *https://make.powerapps.com
 */
const myapp = new UserAgentApplication({
	auth: {
		  clientId: "48a8b1f3-9c77-40d8-a107-fd03bcfa111c",
		  //clientId:"7f8b43c3-4713-4c9e-b980-55947f7a1c3c",
		  //authority: "https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47/", 
		  authority: "https://login.microsoftonline.com/171e6fe3-4b6b-4a50-8303-204ed1bc8339/",
		  redirectUri: "https://" + redirect,
	},
	cache: {
		  cacheLocation: "sessionStorage",
		  storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
	}
  });
const nox:any={
	clientId:"3e6ed93d-966a-4ee3-aab6-b1e6d4fe8314",
	authority:"https://login.microsoftonline.com/171e6fe3-4b6b-4a50-8303-204ed1bc8339/",
};
  const msalConfig:msal2.ClientConfiguration = {
	authOptions:nox,
	
  };



export class pcfMsal implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _out:string;
	private _inputElement: React.ReactElement;
	private _container: HTMLDivElement;
	private _buttonContainer: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private notifyOutputChanged: () => void;
	private _msalConfig:Configuration;
	// Add here the scopes to request when obtaining an access token for MS Graph API
  	// for more, visit https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-core/docs/scopes.md
  	//private _loginRequest = {
	//	scopes: ["openid", "profile", "User.Read"]
  	//};
	private _userAgentApplication:UserAgentApplication;
	private _sessionId:string;
	private _sessionEntityName = "new_sessionaudit";
	private _sessionEntity:ComponentFramework.WebApi.Entity = {
		new_name:"test name"
	};
	private _buttProps:button.IButtonExampleProps = {
				disabled:false,
				checked:false,
				onButtonClicked: this.signIn.bind(this)
			};
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Config object to be passed to Msal on creation.
		// For a full list of msal.js configuration parameters, 
		// visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
		this._container = container;
		context.mode.trackContainerResize(true);
		//this._container = document.createElement("<div>");
		this._buttonContainer = document.createElement("div");
		//styling added here from trial and error, if removed, scroll bars will not show up on the grid list
		this._buttonContainer.classList.add("button-container");
		this._buttonContainer.setAttribute("style","height:99%;width:99%;position: inherit;overflow:scroll");
		this._container.appendChild(this._buttonContainer);
		//container.appendChild(this._container);
		this._context = context;
        this.notifyOutputChanged = notifyOutputChanged;
		if(this._context.parameters.sampleProperty.raw)this._sessionId = this._context.parameters.sampleProperty.raw
		else this._sessionId = "3e6ed93d-966a-4ee3-aab6-b1e6d4fe8314";
		//alert(redirect);
		this._msalConfig = {
			auth: {
					clientId: "48a8b1f3-9c77-40d8-a107-fd03bcfa111c",
					//clientId:"7f8b43c3-4713-4c9e-b980-55947f7a1c3c",
					//authority: "https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47/", 
					authority: "https://login.microsoftonline.com/171e6fe3-4b6b-4a50-8303-204ed1bc8339/",
					redirectUri: "https://" + redirect,
					
				},
		  }; 
		   

	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 * Instantiates ButtonDefaultExample class from reactButton.tsx file
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
		ReactDOM.render(
            this._inputElement = React.createElement(button.ButtonDefaultExample,this._buttProps),
            this._container
        );

	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {output:this._out};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void{}

	/**
	 * bound to button onclick handler for reactButton.tsx file
	 * called when Sign In is clicked in the UI
	 * triggers msal login popup with redirecturi set to current page to allow popup to close seamlessly
	 */
	public signInDevice():void{

		let pca = new DeviceCodeClient(msalConfig);// msal2.PublicClientApplication(msalConfig);
		let deviceCodeRequest:CommonDeviceCodeRequest = {
    		deviceCodeCallback: (response:DeviceCodeResponse) => (alert(response.message)),
			scopes: ["user.read"],
			correlationId: "123",
			authority:"https://login.microsoftonline.com/171e6fe3-4b6b-4a50-8303-204ed1bc8339/",
			timeout: 20
		};
			pca.acquireToken(deviceCodeRequest).then((response)=>(
				console.log(response?.accessToken
					)));
			
	}
	public signInDevice2():void{
		const msalConfig2 = {
			auth: {
				clientId: "3e6ed93d-966a-4ee3-aab6-b1e6d4fe8314",
				authority: "https://login.microsoftonline.com/gronbeckinc.onmicrosoft.com", 
			},
		};

		const pca = new msal3.PublicClientApplication(msalConfig2);

		const deviceCodeRequest = {
			deviceCodeCallback: (response:DeviceCodeResponse) => (alert(response.message)),
			scopes: ["https://gronbeckd365.crm.dynamics.com/user_impersonation"],
			timeout: 20,
		};
		
		pca.acquireTokenByDeviceCode(deviceCodeRequest).then((response) => {
			alert(JSON.stringify(response));
		}).catch((error) => {
			alert(JSON.stringify(error.message));
		});
		
			
	}
	public signIn():void {
		//set redirect uri again so that its up to date
		//if(window.location.host == "localhost:8181")this._msalConfig.auth.redirectUri = "http://localhost:8181";
		//else if(window.location.host == "127.0.0.1:1030") this._msalConfig.auth.redirectUri = "http://127.0.0.1:1030";
		//else this._msalConfig.auth.redirectUri = "https://"+window.location.host;
		this._msalConfig.auth.redirectUri = "https://" + redirect;
		//alert(this._msalConfig.auth.redirectUri);
		let myuseragentapp:UserAgentApplication = new UserAgentApplication(this._msalConfig);
		
		var thisRef = this;
		//myuseragentapp.acquireTokenPopup({scopes:["https://gronbeckinc.crm.dynamics.com/user_impersonation"],prompt:"select_account"}).then(function(loginResponse){
		//https://microsoft.onmicrosoft.com/e1163e19-e5b4-410a-bd0c-dde0657d737b/user_impersonation
		//myuseragentapp.acquireToken
		myuseragentapp.loginPopup({scopes:["https://gronbeckd365.crm.dynamics.com/user_impersonation"],prompt:"select_account"}).then(function(loginResponse){
			myuseragentapp.acquireTokenSilent({scopes:["https://gronbeckd365.crm.dynamics.com/user_impersonation"],prompt:"select_account"}).then(function(loginResponse){
			thisRef._sessionEntity.new_success = true;
			thisRef._sessionEntity.new_name = "Account: " + loginResponse.account.userName + " at: "+ Date.now();

			$.ajaxSetup({
				headers:{
				   	"Authorization": "Bearer "+loginResponse.accessToken,
				   	"Accept": "application/json",
				  	"Content-Type": "application/json; charset=utf-8",
				  	"OData-MaxVersion": "4.0" ,
				   	"OData-Version": "4.0"  
				}
			 });
			// $.get(webAPIURL).then(function(loginResponse){
			$.post(webAPIURL,JSON.stringify(thisRef._sessionEntity)).then(function (loginResponse) {
			//thisRef._context.webAPI.createRecord(thisRef._sessionEntityName,thisRef._sessionEntity )
				alert("Logged Success Login Attempt");
			}).catch(function (error:any) {
				alert(error.message);
				thisRef._out = error.message;
				console.log(error);
			}
			);
		}).catch(function(error){alert(error.message);})}).catch(function(error){
			thisRef._sessionEntity.new_success = false;
			 //$.get(webAPIURL).then(function(loginResponse){
			$.post(webAPIURL,JSON.stringify(thisRef._sessionEntity)).then(function (loginResponse) {
			//thisRef._context.webAPI.createRecord(thisRef._sessionEntityName,thisRef._sessionEntity ).then(function (loginResponse) {
				alert("Logged Success Login Attempt");
			}).catch(function (error:any) {
				alert(error.message);
				thisRef._out = error.message;
				console.log(error);
			}
			);
		});
		

	}
	public logSuccess(ctx:pcfMsal):void{
		ctx._sessionEntity.new_success = true;
		//this._context.webAPI.updateRecord(this._sessionEntityName,this._sessionId,this._sessionEntity ).then(function (loginResponse) {
		//	alert("Logged Success Login Attempt");
		//}).catch(function (error) {
		//	alert(error);
		//	console.log(error);
		//}
		//);
	}
	public logFailure(ctx:pcfMsal):void{
		ctx._sessionEntity.new_success = false;
		//this._sessionEntity.new_success = false;
		//this._context.webAPI.updateRecord(this._sessionEntityName,this._sessionId,this._sessionEntity ).then(function (loginResponse) {
		//	alert("Logged Failed Login Attempt");
		//}).catch(function (error) {
		//	alert(error);
		//	console.log(error);
		//});
	}
	public signOut():void {
		this._userAgentApplication.logout();
	}
}