import React from 'react';
import styles from './MsTeamsClock.module.scss';
import { IMsTeamsClockProps } from './IMsTeamsClockProps';
import { IMsTeamsClockState } from './IMsTeamsClockState';
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/calendars";
import moment from 'moment';
import { IEvents } from '../Services/IEvents';
import { Shimmer} from 'office-ui-fabric-react/lib/Shimmer';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TooltipHost, TooltipDelay, DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import { getId } from 'office-ui-fabric-react/lib/Utilities';






export default class MsTeamsClock extends React.Component<IMsTeamsClockProps, IMsTeamsClockState> {

  constructor(props: IMsTeamsClockProps, state: IMsTeamsClockState) {
    super(props);

    graph.setup({
      spfxContext: this.props.context
    });
    this._indicatorClick = this._indicatorClick.bind(this);
    this.state = {loaded:false, isCalloutVisible: true, items: [],selectedEvent: ["","","","","","","",""]};
  }



  public componentDidMount() {
    this._shimmer();
    this._filterMeetings();

    setInterval(this.drawClock.bind(this), 1000);
  }


  public componentDidUpdate(prevProps, prevState) {
    if (this.state.items !== prevState.items) {
      this._filterMeetings();
    }
    if (this.state.loaded !== prevState.loaded) {
      this._shimmer();
    }
    if (this.state.selectedEvent !== prevState.selectedEvent) {
      this._MeetingLink();
    }

  }
  private _hostId: string = getId('tooltipHost');
  public render(): React.ReactElement<IMsTeamsClockProps> {



    const eventCanceled = this.state.selectedEvent[4];
    const responseStatus = this.state.selectedEvent[6];


    let resMessage = "";
    if (responseStatus === "notResponded" )
    {
      resMessage = "Event Not Accepted";
    }
    else
    {
      resMessage = responseStatus;
    }
    const messageCanceled = eventCanceled ? 'Not Canceled' : 'Canceled';
    let newstr;
    let newstr2 = "";
    var n = false;
    if(this.state.selectedEvent[3] != null && this.state.selectedEvent[3] != 'undefined')
    {
       newstr = this.state.selectedEvent[3];


       n = this.state.selectedEvent[3].includes("Click here to join");
      if (n == true)
      {
        newstr = this.state.selectedEvent[3].replace("Click here to join"," " );
        newstr2 = this.state.selectedEvent[7].joinUrl;
      }
      else
      {
        newstr2 = "#";
      }

    }




    return (
      <div className={styles.msTeamsClock}>
        <div className={styles.title}>{this.props.title}</div><hr></hr>
       <Shimmer  isDataLoaded={this.state.loaded}></Shimmer>
        <div className={styles.container} id="container">
            <div className={styles.col1}>
              <ul>
                <li>
                <TooltipHost
          calloutProps={{ gapSpace: 20 }}
          tooltipProps={{
            onRenderContent: () => {
              return (
                <div style={{fontSize:"14px",backgroundColor:"slateblue", color:"white",borderBlockColor:"black"}}>
                  {this.props.description}
                </div>
              );
            }
          }}
          delay={TooltipDelay.zero}
          id={this._hostId}
          directionalHint={DirectionalHint.bottomCenter}
        >
         <DefaultButton  aria-labelledby={this._hostId} text="Info" />

        </TooltipHost>
                </li>
                <li className={styles.leg1}>No Events:</li>
                <li className={styles.leg2}>Meeting:</li>
                <li className={styles.leg4}>No Response:</li>
                <li className={styles.leg3}>Break:</li>
              </ul>
            </div>

            <div className={styles.col2}>
              <div className={styles.clock}>
                <div className={styles.centernut}>
                </div>
                <div className={styles.centernut2}>
                </div>
                <div className={styles.indicators} id="indicators">

                  <div id="12:30" style={{ cursor: 'default',pointerEvents: 'none'}} aria-disabled="false" onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="1:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="1:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="2:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="2:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="3:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="3:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="4:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="4:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="5:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="5:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="6:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="6:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="7:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="7:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="8:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="8:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="9:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="9:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="10:0" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="10:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="11:0" style={{ cursor: 'default', pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="11:30" style={{ cursor: 'default',pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                  <div id="12:0" style={{ cursor: 'default', pointerEvents: 'none'}} onClick={this._indicatorClick.bind(this)}>
                  </div>
                </div>
                <div className={styles.sechand}>
                  <div className={styles.sec} id="sec">
                  </div>
                </div>
                <div className={styles.minhand}>
                  <div className={styles.min} id="min">
                  </div>
                </div>
                <div className={styles.hrhand}>
                  <div className={styles.hr} id="hr">
                  </div>
                </div>
              </div>
            </div>

            <div className={styles.col3}>
              <div id="info" className={styles.info} style={{ display: 'none' }}>
                <span className={styles.close} onClick={e => this._CloseModal(this)}>&times;</span>

                <><p><b>Subject: </b>{this.state.selectedEvent[0]}</p>
                  <p><b>Start Date: </b>{moment(this.state.selectedEvent[1]).format('llll')}</p>
                  <p><b>End Date: </b>{moment(this.state.selectedEvent[2]).format('llll')}</p>
                  <p><b>Response Status: </b>{resMessage}</p>
                  <p><b>Event Status: </b>{messageCanceled}</p>
                  <p><b>Location: </b>{this.state.selectedEvent[5]}</p>
                  <p><b>Intro: </b>{newstr}  <a style={{ display: 'block', backgroundColor:"white"}} id={this.state.selectedEvent[1]} href={newstr2} target="_blank" rel="noopener noreferrer"> Click here to join the meeting</a></p>
                </>
              </div>
            </div>
          </div>
      </div>
    );
  }

private _MeetingLink()
{
  var elem = document.getElementById('info').querySelectorAll("a");
    for (var i = 0; i < elem.length; i++) {
     var lnk = elem[i].getAttribute("href");
      if (lnk == "#" && lnk.length > 0 ) {
        if (document.getElementById(elem[i].id)) {
          var x = (document.getElementById(elem[i].id) as HTMLLinkElement);
          x.style.display = "none";
        }
      }
      else if(lnk != "#" && lnk.length > 0 ) {
        if (document.getElementById(elem[i].id)) {
          var y = (document.getElementById(elem[i].id) as HTMLLinkElement);
          y.style.display = "block";
        }
      }
    }

}
private _shimmer()
{
  var info  = document.getElementById('container');
  if (this.state.loaded == false)
  {

    info.style.display = 'none';
  }
  else
  {
    info.style.display = '';
  }
}

  private _CloseModal = (e): void => {
    var info  = document.getElementById('info');
    info.style.display = 'none';
  }


  public drawClock() {
    const sec = document.getElementById("sec");
    const min = document.getElementById("min");
    const hr = document.getElementById("hr");

    let time = new Date();
    let secs = time.getSeconds() * 6;
    let mins = time.getMinutes() * 6;
    let hrs = time.getHours() * 30;
    sec.style.transform = `rotateZ(${secs}deg)`;
    min.style.transform = `rotateZ(${mins}deg)`;
    hr.style.transform = `rotateZ(${hrs + (mins / 12)}deg)`;
  }
  public async _filterMeetings() {
    this.setState({loaded:false});
    var tablinks;
    tablinks = document.getElementById('indicators').getElementsByTagName("div");
    for (var j = 0; j < tablinks.length; j++) {
      tablinks[j].style.background = "slateblue";
      tablinks[j].style.cursor = "default";
      tablinks[j].style.pointerEvents = "none";
    }
    var items: Array<IEvents> = new Array<IEvents>();
    var today = moment();
    var day = today;
    const startOfDay = moment(day).startOf("day");
    const endOfDay = moment(day).endOf("day");



      const caleEvents = await graph.me.calendar
      .calendarView(startOfDay.toISOString(), endOfDay.toISOString())
      .select('subject', 'start', 'end','location','bodyPreview','isCancelled','responseStatus','onlineMeeting')();


    caleEvents.map((item: any) => {


      var startDateUtc = moment.utc(item.start.dateTime);
      var StartDt = startDateUtc.local();
      var endDateUtc = moment.utc(item.end.dateTime);
      var EndDt = endDateUtc.local();

      items.push({
        Subject: item.subject,
        Start: StartDt,
        End: EndDt,
        TimeZone:item.start.timeZone,
        Location:item.location.displayName,
        BodyPreview:item.bodyPreview,
        isCancelled:item.isCancelled.toString(),
        responseStatus:item.responseStatus.response,
        onlineMeeting:item.onlineMeeting
      });
     });

    this.setState({ items: items });



    var eventTime  = "";
    for (let index = 0; index < this.state.items.length; index++) {
      eventTime = moment(this.state.items[index].Start).format("h:m").toString();
      for (var i = 0; i < tablinks.length; i++) {
        if (tablinks[i].id === eventTime)
        {
         if(this.state.items[index].Subject === "Break" || this.state.items[index].Subject === "Lunch Break")
         {
          tablinks[i].style.background = "orange";
          tablinks[i].style.cursor = "pointer";
          tablinks[i].style.pointerEvents = "auto";
         }
         else
         {
          if(this.state.items[index].responseStatus === "notResponded")
          {
          tablinks[i].style.background = "red";
          tablinks[i].style.cursor = "pointer";
          tablinks[i].style.pointerEvents = "auto";
          }
          else
          {
          tablinks[i].style.background = "black";
          tablinks[i].style.cursor = "pointer";
          tablinks[i].style.pointerEvents = "auto";
          }
         }
        }
      }
    }
    this.setState({loaded:true});
  }

  public _indicatorClick(event) {
    this.setState({selectedEvent: ["No events in your calendar for the selected time!"] });
    var eventTime  = "";

    for (let index = 0; index < this.state.items.length; index++) {
      eventTime = moment(this.state.items[index].Start).format("h:m").toString();
      if (eventTime === event.target.id)
       {

        this.setState({
          selectedEvent: [
            this.state.items[index].Subject,
            this.state.items[index].Start,
            this.state.items[index].End,
            this.state.items[index].BodyPreview,
            this.state.items[index].isCancelled,
            this.state.items[index].Location,
            this.state.items[index].responseStatus,
            this.state.items[index].onlineMeeting,
          ]
        });

      }

    }

    var info  = document.getElementById('info');
    info.style.display = 'block';
  }
}



