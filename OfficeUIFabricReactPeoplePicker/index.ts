import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { PeoplePickerTypes, IPeopleProps, IPeoplePersona } from './Peoplepicker';

export class OfficeUIFabricReactPeoplePickerPCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private theContainer: HTMLDivElement;
	private notifyOutputChanged: () => void;
	private _context: ComponentFramework.Context<IInputs>;
	private props: IPeopleProps = {
		//tableValue: this.numberFacesChanged.bind(this),
		peopleList: this.peopleList.bind(this),
	}

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		// Add control initialization code
		this.notifyOutputChanged = notifyOutputChanged;
		this._context = context;
		this.theContainer = container;
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public async updateView(context: ComponentFramework.Context<IInputs>) {
		// Add code to update control view
		this.props.context = context;
		const myProps = this.props;
		if (context.parameters.peoplePicker.raw !== null) {
			if (context.parameters.peoplePicker.raw!.indexOf("text") > 1) {
				this.props.preselectedpeople = JSON.parse(context.parameters.peoplePicker.raw!);
			}
		}
		ReactDOM.unmountComponentAtNode(this.theContainer);
		ReactDOM.render(
			React.createElement(
				PeoplePickerTypes,
				myProps
			),
			this.theContainer
		);

	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			peoplePicker: JSON.stringify(this.props.people)
		};
	}

	private peopleList(newValue: IPeoplePersona[]) {
		if (this.props.people !== newValue) {
			if (newValue.length === 0) {
				this.props.people = undefined;
			}
			else {
				this.props.people = newValue;
			}

			this.notifyOutputChanged();
		}
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
		ReactDOM.unmountComponentAtNode(this.theContainer);
	}
}