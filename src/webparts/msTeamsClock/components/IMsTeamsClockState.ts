import { IEvents } from './../Services/IEvents';
export interface IMsTeamsClockState {
  items:IEvents[];
  selectedEvent:any[];
  isCalloutVisible: boolean;
  loaded:boolean;
}
