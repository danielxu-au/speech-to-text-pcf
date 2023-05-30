import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as SpeechSDK from 'microsoft-cognitiveservices-speech-sdk';
import { AutoDetectSourceLanguageConfig } from "microsoft-cognitiveservices-speech-sdk";

export class SpeechToText implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _notifyOutputChanged: () => void;

    // component attributes
	private _context : ComponentFramework.Context<IInputs>;
	private _container : HTMLDivElement;
    private _buttonDiv : HTMLDivElement;
    private _isInitiated : boolean = false;
    private _isInListenMode : boolean = false;

    // component attributes
    private _subscriptionKey : string;
    private _region : string;
    private _sourceLanguage : string;
    private _targetLanguage : string;
    private _buttonMicColor : string | undefined = "#ff0000";
    private _buttonStopColor : string | undefined = '#00ff00';
    private _autoDetect : boolean = false;
    private _translationsJson : string;

    // output attributes
    private _state : string = "waiting"; // waiting, listening, finished
    private _originalText : string = "";
    private _translatedText : string = "";
    private _errorText : string = "";

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
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // save the context
        this._context = context;
        this._context.mode.trackContainerResize(true);

        // save the notifyOutputChanged
        this._notifyOutputChanged = notifyOutputChanged;

        // Add control initialization code
        this._container = container;

        // set default values
        this._state = "waiting";
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        this.updateStateFromContext(context);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {
            "state" : this._state,
            "spokenText" : this._originalText,
            "translatedText" : this._translatedText,
            "errorText" : this._errorText,
            "translationsJSON" : this._translationsJson
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }

    /**
     *  Custom functions
     */
    public updateStateFromContext(context: ComponentFramework.Context<IInputs>) : void
    {
        // Add code to update control view
        this._subscriptionKey = context.parameters.subscriptionKey.raw as string;
        this._region = context.parameters.region.raw as string;
        this._sourceLanguage = context.parameters.sourceLanguage.raw as string;
        this._targetLanguage = context.parameters.targetLanguage.raw as string;
        this._state = context.parameters.state.raw as string;

        const buttonMicColor = context.parameters.micButtonColor.raw as string;
        const buttonStopColor = context.parameters.stopButtonColor.raw as string;
        const autoDetect = context.parameters.autoDect.raw as boolean;

        if(buttonMicColor != undefined || buttonMicColor != "" || buttonMicColor != null)
        {
            this._buttonMicColor = buttonMicColor;
        }

        if(buttonStopColor != undefined || buttonStopColor != null || buttonStopColor != "")
        {
            this._buttonStopColor = buttonStopColor;
        }

        if(autoDetect != undefined || autoDetect != null)
        {
            this._autoDetect = autoDetect;
        }
        
        console.log(`subscription key is ${this._subscriptionKey}`);
        console.log(`region is ${this._region}`);
        console.log(`spoken language is ${this._sourceLanguage}`);
        console.log(`translated language is ${this._targetLanguage}`);
        console.log(`autoDetect is ${this._autoDetect}`);
        console.log(`mic button color is ${this._buttonMicColor}`);
        console.log(`stop button color is ${this._buttonStopColor}`);


        if(!this._isInitiated)
        {
            // create the translation div & button
            this._buttonDiv = document.createElement("div");
            this._buttonDiv.id = `button-div`;
            this._buttonDiv.className = `button-div`;
            this._buttonDiv.style.width = `100%`;
            this._buttonDiv.style.height = `100%`;
            this._buttonDiv.style.cursor = `pointer`; 
            this._buttonDiv.innerHTML = `<svg width="${this._context.mode.allocatedWidth}" height="${this._context.mode.allocatedHeight}" fill="none" viewBox="0,0,1024,1024" xmlns="http://www.w3.org/2000/svg"><g clip-path="url(#prefix__clip0_239_2)"><circle cx="512" cy="512" r="448" fill="${this._buttonMicColor}"/><circle cx="512" cy="512" r="480" stroke="${this._buttonMicColor}" stroke-opacity=".5" stroke-width="64"/><rect x="388.678" y="256" width="243.45" height="364.551" rx="121.725" fill="#fff"/><path d="M694.551 499.658c0 100.668-81.607 182.276-182.275 182.276S330 600.326 330 499.658" stroke="#fff" stroke-width="64" stroke-linecap="round"/><path d="M544.276 707.32v-32h-64v32h64zm-64 60.68c0 17.673 14.327 32 32 32 17.673 0 32-14.327 32-32h-64zm0-60.68V768h64v-60.68h-64z" fill="#fff"/></g><defs><clipPath id="prefix__clip0_239_2"><path fill="#fff" d="M0 0h1024v1024H0z"/></clipPath></defs></svg>`;
            this._buttonDiv.addEventListener('click', this.startListening.bind(this));
            this._container.appendChild(this._buttonDiv);  

            // set the initialised state to true
            this._isInitiated = true;
        } else {
            if(this._isInListenMode) 
            {
                this.startListeningUpdateUIComponents();
            } 
            else 
            {
                this.stopListeningUpdateUIComponents();
            }
        }
    } 


    public startListening() : void
    {
        // state
        console.log('I am listening ...');

        // reset the text values
        this._translatedText = "";
        this._originalText = "";

        // create the speech recogniser
        var speechConfig = SpeechSDK.SpeechTranslationConfig.fromSubscription(this._subscriptionKey, this._region);
        speechConfig.speechRecognitionLanguage = this._sourceLanguage;

        const targetLanguages = ["zh-Hans","cy","de","en","fr","ga","es","it","nl","pt-pt","ru","sv"];

        targetLanguages.forEach((language)=>{
            speechConfig.addTargetLanguage(language);
        });
    
        let audioConfig  = SpeechSDK.AudioConfig.fromDefaultMicrophoneInput();

        this.startListeningUpdateUIComponents();
        this._notifyOutputChanged();

        if(this._autoDetect) 
        {
            var autoDetectSourceLanguageConfig = AutoDetectSourceLanguageConfig.fromLanguages(["fr-FR","en-GB","de-DE","es-ES"]);
            var speechRecognizer = SpeechSDK.SpeechRecognizer.FromConfig(speechConfig, autoDetectSourceLanguageConfig, audioConfig);
            speechRecognizer.recognizeOnceAsync(
                (result) => {
                    console.log(`result is ${JSON.stringify(result)}`);
                    if (result.reason === SpeechSDK.ResultReason.TranslatedSpeech) 
                    {
                        //let original = result.text;
                        //this._originalText += original;
                        //let translation = result.translations.get(this._targetLanguage);
                        //this._translatedText += translation;
                        //console.log(`done listening - translated text is ${this._translatedText}`);
                    }
                    speechRecognizer.close();
                    this.stopListeningUpdateUIComponents();
                    this._notifyOutputChanged();
                },
                (err) => {
                    this._errorText += err;
                    console.log(`error: ${this._errorText}`);
                    speechRecognizer.close();
                    this.stopListeningUpdateUIComponents();  
                    this._notifyOutputChanged();
                }
            )

        }
        else
        {
            var translationRecognizer = new SpeechSDK.TranslationRecognizer(speechConfig, audioConfig);
            translationRecognizer.recognizeOnceAsync(
                (result) => {
                     console.log(`result is ${JSON.stringify(result)}`);
                     if (result.reason === SpeechSDK.ResultReason.TranslatedSpeech) 
                     {
                        let original = result.text;
                        this._originalText += original;
                        let translation = result.translations.get(this._targetLanguage);
                        this._translatedText += translation;
                        this._translationsJson = result.json;
                        console.log(`done listening - translated text is ${this._translatedText}`);
                        console.log(`${result.json}`);
                     }
                     translationRecognizer.close();
                     this.stopListeningUpdateUIComponents();
                     this._notifyOutputChanged();
                 },
                 (err) => {
                     this._errorText += err;
                     console.log(`error: ${this._errorText}`);
                     translationRecognizer.close();
                     this.stopListeningUpdateUIComponents();  
                     this._notifyOutputChanged();
                 }
             );
        }
    }

    public startListeningUpdateUIComponents() 
    {
        this._buttonDiv.innerHTML=`<svg width="${this._context.mode.allocatedWidth}" height="${this._context.mode.allocatedHeight}" viewBox="0,0,1024,1024" fill="none" xmlns="http://www.w3.org/2000/svg"><g clip-path="url(#prefix__clip0_236_16)"><circle cx="512" cy="512" r="448" fill="${this._buttonStopColor}"/><circle cx="512" cy="512" r="480" stroke="${this._buttonStopColor}" stroke-opacity=".5" stroke-width="64"/><rect x="256" y="256" width="512" height="512" rx="64" fill="#fff"/></g><defs><clipPath id="prefix__clip0_236_16"><path fill="#fff" d="M0 0h1024v1024H0z"/></clipPath></defs></svg>`
        this._state = "listening";
        this._originalText = "";
        this._translatedText = "";
        this._isInListenMode = true;
    }

    public stopListeningUpdateUIComponents() 
    {
        this._buttonDiv.innerHTML = `<svg width="${this._context.mode.allocatedWidth}" height="${this._context.mode.allocatedHeight}" fill="none" viewBox="0,0,1024,1024" xmlns="http://www.w3.org/2000/svg"><g clip-path="url(#prefix__clip0_239_2)"><circle cx="512" cy="512" r="448" fill="${this._buttonMicColor}"/><circle cx="512" cy="512" r="480" stroke="${this._buttonMicColor}" stroke-opacity=".5" stroke-width="64"/><rect x="388.678" y="256" width="243.45" height="364.551" rx="121.725" fill="#fff"/><path d="M694.551 499.658c0 100.668-81.607 182.276-182.275 182.276S330 600.326 330 499.658" stroke="#fff" stroke-width="64" stroke-linecap="round"/><path d="M544.276 707.32v-32h-64v32h64zm-64 60.68c0 17.673 14.327 32 32 32 17.673 0 32-14.327 32-32h-64zm0-60.68V768h64v-60.68h-64z" fill="#fff"/></g><defs><clipPath id="prefix__clip0_239_2"><path fill="#fff" d="M0 0h1024v1024H0z"/></clipPath></defs></svg>`;
        this._state = "finished";
        this._isInListenMode = false;
    }
}
