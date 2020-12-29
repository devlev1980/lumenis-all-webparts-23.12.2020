import {Environment, EnvironmentType, Version} from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneCheckbox,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import {escape} from '@microsoft/sp-lodash-subset';

import * as strings from 'LumenisAllWebPartStrings';
 require('./LumenisAllWebPart.module.scss');
require('./all_webparts.module.scss');
require('./style.css');



import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import * as pnp from 'sp-pnp-js';
import {Web} from 'sp-pnp-js';
import * as $ from 'jquery';
import * as moment from 'moment';

export interface ILumenisAllWebPartProps {
//  Header
  mainBannerImg: string;
// Usefull Links
  description_UsefulLinks: string;
  wpTitle_UsefulLinks: string;
  wpWebUrl_UsefulLinks: string;
  listName_UsefulLinks: string;
  // QuickLinks Webpart
  Link1: string;
  Link1Text: string;
  LinkImage1: string;
  Link2: string;
  Link2Text: string;
  LinkImage2: string;
  Link3: string;
  Link3Text: string;
  LinkImage3: string;
  Link4: string;
  Link4Text: string;
  LinkImage4: string;
  Link5: string;
  Link5Text: string;
  LinkImage5: string;
  Link6: string;
  Link6Text: string;
  LinkImage6: string;
  Link7: string;
  Link7Text: string;
  LinkImage7: string;
  Link8: string;
  Link8Text: string;
  LinkImage8: string;
  Link9: string;
  Link9Text: string;
  LinkImage9: string;
  // Greetings
  descriptionGreetings: string;
  listNameGreetings: string;
  webUrlGreetings: string;
  wpTitleGreetings: string;
  sendYourGreetings: string;
  separator: string;
  daysBefore: string;
  daysAfter: string;
  displayDate: boolean;
  // Photo gallery
  photoGalleryTitle: string;
  linksImagesTitle: string;
  description: string;
  photosLibraryName: string;
  Link1PhotoGallery: string;
  Link1TextPhotoGallery: string;
  LinkImage1PhotoGallery: string;
  Link2PhotoGallery: string;
  Link2TextPhotoGallery: string;
  //LinkImage2: string;
  Link3PhotoGallery: string;
  Link3TextPhotoGallery: string;
  //LinkImage3: string;
  Link4PhotoGallery: string;
  Link4TextPhotoGallery: string;
  //LinkImage4: string;
}

export interface ISPLists {
  Files: ISPList[];
}

export interface ISPList {
  ServerRelativeUrl: string;
  ListItemAllFields: ISPField;
}

export interface ISPField {
  Title: string;
  Date: string;
  ImageTitle: string;
  Link: string;
}

export default class LumenisAllWebPart extends BaseClientSideWebPart<ILumenisAllWebPartProps> {
  private _newEmployeesListName = 'NewEmployees';
  private slideIndex: number = 1;

  public render(): void {
    let d = new Date();
    let thisTime = d.getTime();
    //region All except greetings
    let NumberOfActiveLinks = 0;
    if (this.properties.Link1 != '') {
      NumberOfActiveLinks += 1;
    }
    if (this.properties.Link2 != '') {
      NumberOfActiveLinks += 1;
    }
    if (this.properties.Link3 != '') {
      NumberOfActiveLinks += 1;
    }
    if (this.properties.Link4 != '') {
      NumberOfActiveLinks += 1;
    }
    if (this.properties.Link5 != '') {
      NumberOfActiveLinks += 1;
    }
    if (this.properties.Link6 != '') {
      NumberOfActiveLinks += 1;
    }
    if (this.properties.Link7 != '') {
      NumberOfActiveLinks += 1;
    }
    if (this.properties.Link8 != '') {
      NumberOfActiveLinks += 1;
    }
    if (NumberOfActiveLinks != 0 && NumberOfActiveLinks != 8) {
      NumberOfActiveLinks = (100 - 12.5 * NumberOfActiveLinks) / 2;
    } else NumberOfActiveLinks = 0;
    let html = ``;
    html += `
 <header style="width: 100%">
 <a href="https://lumenis.sharepoint.com/sites/Portal">
<!--         <img src="https://lumenis.sharepoint.com/_api/v2.1/drives/b!gTFHrKKy30-wJdR8QTo71wK3BPOGa6BOtVvcKCVKTCtS1FB9GvW6TJT5gsqaxIr6/items/01K54NSDGIFZHKI25RDVCY2GZWSR2VHFIS/thumbnails/0/c1600x99999/content?preferNoRedirect=true&prefer=extendCacheMaxAge&clientType=modernWebPart" alt="">-->
         <img src="${escape(this.properties.mainBannerImg)}" alt="">

</a>
</header>
        <div class="all-webparts__container" style="display: grid;
        grid-template-columns: 18vw 43vw 18vw;height: auto; grid-gap: 30px;margin: 0 9vw">

        <div class="left__sidebar">
             <div class="sidebar_container" id="usefulLinksWP">
                <div style="background: #fff;border-radius: 10px;height: 84%"><h3 id="usefulLinksWPTitle">${this.properties.wpTitle_UsefulLinks}</h3>
                 <div class="useful_list"></div></div>
                 <div class="sahar__btn">
             <button>
             <a href="https://lumenis.sharepoint.com/sites/Portal">
                         <img src="https://lumenis.sharepoint.com/sites/Portal/_layouts/15/getpreview.ashx?resolution=3&guidSite=ac473181b2a24fdfb025d47c413a3bd7&guidWeb=f304b7026b864ea0b55bdc28254a4c2b&guidFile=eba47dd82d884f7a82084cd016d381b5&clientType=modernWebPart" alt="">

</a>
</button>
</div>
             </div>

         </div> `;


    // Quick Links
    if (this.properties.Link1 != '') {
      html += `
<div class="main" >
  <div class="LinksWrapper">
                <h1 class="title">מידע לעובדים</h1>
             </div>
       <div class="LinksWrapper-Raw">
         <div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link1)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage1)}"></a>
          </div>
          <div class="LinkText">
            <a href="${escape(this.properties.Link1)}">${escape(this.properties.Link1Text)}</a>
          </div>
</div>`;
    }
    if (this.properties.Link2 != '') {
      html += `<div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link2)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage2)}"></a>
          </div>
          <div class="LinkText" style="position: relative;left: -18px">
            <a href="${escape(this.properties.Link2)}">${escape(this.properties.Link2Text)}</a>
          </div>
        </div>`;
    }
    if (this.properties.Link3 != '') {
      html += `<div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link3)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage3)}"></a>
          </div>
          <div class="LinkText">
            <a href="${escape(this.properties.Link3)}">${escape(this.properties.Link3Text)}</a>
          </div>
        </div>`;
    }
    if (this.properties.Link4 != '') {
      html += `<div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link4)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage4)}"></a>
          </div>
          <div class="LinkText">
            <a href="${escape(this.properties.Link4)}">${escape(this.properties.Link4Text)}</a>
          </div>
        </div>`;
    }
    if (this.properties.Link5 != '') {
      html += `<div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link5)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage5)}"></a>
          </div>
          <div class="LinkText">
            <a href="${escape(this.properties.Link5)}">${escape(this.properties.Link5Text)}</a>
          </div>
        </div>`;
    }
    if (this.properties.Link6 != '') {
      html += `<div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link6)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage6)}"></a>
          </div>
          <div class="LinkText">
            <a href="${escape(this.properties.Link6)}">${escape(this.properties.Link6Text)}</a>
          </div>
        </div>`;
    }
    if (this.properties.Link7 != '') {
      html += `<div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link7)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage7)}"></a>
          </div>
          <div class="LinkText">
            <a href="${escape(this.properties.Link7)}">${escape(this.properties.Link7Text)}</a>
          </div>
        </div>`;
    }
    if (this.properties.Link8 != '') {
      html += `<div class="LinksTab">
          <div class="LinkImage">
            <a href="${escape(this.properties.Link8)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage8)}"></a>
          </div>
          <div class="LinkText">
            <a href="${escape(this.properties.Link8)}">${escape(this.properties.Link8Text)}</a>
          </div>
        </div>`;
    }


    // let html2 = '<h1>aaa</h1>'
    html += `<div class="LinksTab">
        <div class="LinkImage">
          <a href="${escape(this.properties.Link9)}"><img class="ImageInLink" src="${escape(this.properties.LinkImage9)}"></a>
        </div>
        <div class="LinkText">
          <a href="${escape(this.properties.Link9)}">${escape(this.properties.Link9Text)}</a>
        </div>


      </div></div>
            `;
    {
      html += `<div id="PhotoGalleryWebpartWrapper">
                 <div id="LinksImages" style="direction: rtl">
<div class="LinksWrapper2">
                     <div class="LinksWrapper-Raw2" ><h2  style="color: #040507";>${this.properties.linksImagesTitle}</h2><div class="raws-wrapper">`;
      if (this.properties.Link1PhotoGallery != '') {
        html += `<div class="LinksTab2" >
                                    <div class="LinkImage2" >
                                      <a href="${escape(this.properties.Link1PhotoGallery)}">
                                      <img class="ImageInLink2" width="100%"   src="${escape(this.properties.LinkImage1PhotoGallery)}">
                                      </a>
                                      <br>
                                    </div>
                                  </div> `;
      }
      if (this.properties.Link2PhotoGallery != '') {
        html += `<div class="LinksTab2">
                                    <div class="LinkImage2">
                                      <button type="button" class="ImageInLink2" onclick="location.href='${this.properties.Link2PhotoGallery}';">${this.properties.Link2TextPhotoGallery}</button>
                                    </div>
                                  </div>`;
        // <a href="${escape(this.properties.Link2)}"><img class="ImageInLink2" style="width: 11.3vw; height: 1.58vw;" src="${escape(this.properties.LinkImage2)}"></a><br>
      }
      if (this.properties.Link3PhotoGallery != '') {
        html += `<div class="LinksTab2">
                                    <div class="LinkImage2">
                                      <button type="button" class="ImageInLink2" onclick="location.href='${this.properties.Link3PhotoGallery}';">${this.properties.Link3TextPhotoGallery}</button>
                                    </div>
                                  </div>`;
        // <a href="${escape(this.properties.Link3)}"><img class="ImageInLink2" style="width: 11.3vw; height: 1.58vw;" src="${escape(this.properties.LinkImage3)}"></a><br>
      }
      if (this.properties.Link4PhotoGallery != '') {
        html += `<div class="LinksTab2">
                                    <div class="LinkImage2">
                                      <button type="button" class="ImageInLink2" onclick="location.href='${this.properties.Link4PhotoGallery}';">${this.properties.Link4TextPhotoGallery}</button>
                                    </div>
                                  </div> </div>`;
        // <a href="${escape(this.properties.Link4)}"><img class="ImageInLink2" style="width: 11.3vw; height: 1.58vw;" src="${escape(this.properties.LinkImage4)}"></a><br>

        html += `</div>

                                   <div id="PhotoGalleryTable">
                                   <div id="Title">
                                   <h2  style="color: #040507">${this.properties.photoGalleryTitle}</h2>
                                    </div>
                                     <div style="width:100%;height: 227px;;  background: #fff;  padding: 16px;  margin: 1.2rem auto;">
                                     <div id="FileProperties">
                                    </div>
                                    <div id="slidesShow-Container">
                                     </div>
                                    <div id="after-slidesShow-Container">
                                     </div>
                                    </div>
                                   </div>

                                  </div
                                </div>


      </div>
      </div>
      </div>
<!--Right sidebar-->
 `;
        {
          html += `<div class="right__sidebar">
    <p id="LumenisGreetingsWpWebPartID" class="anchorinpage"></p>
      <div class="event_wrapper" id="greetingsWP">
    <h3 class="WPtitle">${escape(this.properties.wpTitleGreetings)}</h3>
      <div class=" WPevents" id="WPevents${thisTime}">
    <div class="scrollEvents" id="scrollEvents${thisTime}"></div>

      </div>
      </div>
 </div>`;
          // tslint:disable-next-line:no-unused-expression
        }`
`;
      }
      this._renderListAsync();
      // let NumberOfActiveLinks = 0;
      if (this.properties.Link1PhotoGallery != '') {
        NumberOfActiveLinks += 1;
      }
      if (this.properties.Link2PhotoGallery != '') {
        NumberOfActiveLinks += 1;
      }
      if (this.properties.Link3PhotoGallery != '') {
        NumberOfActiveLinks += 1;
      }
      if (this.properties.Link4PhotoGallery != '') {
        NumberOfActiveLinks += 1;
      }

      if (NumberOfActiveLinks != 0 && NumberOfActiveLinks != 8) {
        NumberOfActiveLinks = (100 - 12.5 * NumberOfActiveLinks) / 2;
      } else NumberOfActiveLinks = 0;

      // let MarginLeft = NumberOfActiveLinks.toString() + '%';
    }
    // tslint:disable-next-line:no-unused-expression
    `
  `;


// <!--// html +=--> `
//     this.domElement.innerHTML = html;


//============================================== Greetings Logic ===================================================
    let greetingsList = this.properties.listNameGreetings;
    let query = this.Build(greetingsList);
    let webUrl = this.properties.webUrlGreetings;
    let sendYourGreetings = this.properties.sendYourGreetings;
    let separator = this.properties.separator;
    let absUrl = this.context.pageContext.site.absoluteUrl;
    let web = pnp.sp.web;

    if (webUrl != '') {
      web = new Web(absUrl + webUrl);
    } else {
      web = new Web(absUrl);
    }
    web.get().then((w) => {
      const q: any = {
        ViewXml: query
      };
      web.lists.getByTitle(greetingsList).getItemsByCAMLQuery(q).then((r: any[]) => {
        var _greetingsList = greetingsList;
        r.forEach((result) => {
          let _listName = _greetingsList;
          let _isNewEmployee = _greetingsList == this._newEmployeesListName;
          let itemTitle = result.Title;
          let itemRole;
          let itemEventType = result.EventType;
          let itemEventAuthorId = result.EventAuthorId;
          let itemMonth = result.EventMonth;
          let itemDay = result.EventDay;
          let babyGender = result.BabyGender;
          if (itemEventType == null) {
            itemEventType = 'birthday';
          }
          let mainID = result.ID;

          let userEmail = '#';
          let userPosition = '';

          let eventImg = '';
          let eventArrowImg = '';
          switch (itemEventType) {
            case 'birthday':
              eventImg = 'birthdayIcon.png';
              itemRole = `${itemDay}${separator}${itemMonth}  מזל טוב`;
              eventArrowImg = 'pinkArrow.jpg';
              break;
            case 'newborn':
              eventImg = 'handshakeIcon.png';
              itemRole = 'ברוך הבא ללומניס';
              eventArrowImg = 'darkBlueArrow.jpg';
              break;
            case 'wedding':
              eventImg = 'weddingIcon.png';
              itemRole = 'מזל טוב לנישואיך';
              eventArrowImg = 'yellowArrow.jpg';
              break;
            case 'baby':
              eventImg = 'strollerIcon.jpeg';
              eventArrowImg = 'lightBlueArrow.jpg';
              switch (babyGender) {
                case 'בן':
                  itemRole = 'מזל טוב להולדת הבן';
                  break;
                case 'בת':
                  itemRole = 'מזל טוב להולדת הבת';
                  break;
              }
              break;
          }

          let selectContainer: Element = this.domElement.querySelector(`#scrollEvents${thisTime}`);
          if (this.properties.displayDate) {
            selectContainer.insertAdjacentHTML('beforeend',
              `<div class="grtmpitem eventItemNews ${itemEventType}" userId="${itemEventAuthorId}" useremail="" id=${mainID}>

                <div class="celebrant_detailes">
                    <div class="evt_ic">
                        <img src="/Style Library/IMF.O365.Lumenis/img/${eventImg}">
                    </div>
                    <div class="info">
                    <div class="user_name"></div>
                     <span class="user_office">${itemRole}</span>

</div>
<!--<div class="eventDate">-->
                    <div class="evt_ic">
                        <img src="/Style Library/IMF.O365.Lumenis/img/${eventArrowImg}">
<!--                   <div> <img src="/Style Library/IMF.O365.Lumenis/img/${eventImg}"></div>-->
                     <a class="wish_btn" href="" title=""><br>${sendYourGreetings}</a>
                    </div>
                </div>
                </div>

                </div>`);
          } else {
            selectContainer.insertAdjacentHTML('beforeend',
              `<div class="grtmpitem eventItemNews ${itemEventType}" userId="${itemEventAuthorId}" useremail="" id=${mainID}>

                    <div class="celebrant_detailes">
                        <div class="user_name"></div>
                        <div class="user_office">${itemRole}</div>
                        <a class="wish_btn" href="" title="">${sendYourGreetings}</a>
                    </div>
                    <div class="eventDate">
                        <div class="evt_ic"><img src="/Style Library/IMF.O365.Lumenis/img/${eventImg}"></div>
                    </div>
                    </div>`);

          }
          var _picture = _isNewEmployee && result.Picture ? result.Picture : undefined;

          if (itemEventAuthorId) {
            web.siteUsers.getById(itemEventAuthorId).get().then((response) => {
              ((res, picture) => {
                // console.log(mainID);
                // console.log(result);
                $(`#WPevents${thisTime} .eventItemNews#${mainID}`).attr('useremail', res.Email);
                $(`#WPevents${thisTime} .eventItemNews#${mainID} .user_pic_date img`).attr('src', picture && picture.Url ? picture.Url : '/_vti_bin/DelveApi.ashx/people/profileimage?userId=' + result.Email);
                $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .user_name`).text(res.Title);
                $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .wish_btn`).attr('href', 'mailto:' + res.Email);
              })(response, _picture);
            });
          } else if (_isNewEmployee && _picture && _picture.Url) {
            $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .user_name`).text(itemTitle);
            $(`#WPevents${thisTime} .eventItemNews#${mainID} .user_pic_date img`).attr('src', _picture.Url);
            $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .wish_btn`).css('visibility', 'hidden');
          }
        });
      });
    }).catch((err) => {
      console.log(err);
    });
    this.domElement.innerHTML = html;
    this.renderLinks();
  }

// ===========================================Useful Links Logic=======================================================
  private renderLinks() {
    let webUrl = this.properties.wpWebUrl_UsefulLinks;
    let usefulLinksList = this.properties.listName_UsefulLinks;
    let absUrl = this.context.pageContext.site.absoluteUrl;
    let web = pnp.sp.web;
    if (webUrl != '') {
      web = new Web(absUrl + webUrl);
    } else {
      web = new Web(absUrl);
    }

    let resultContainer: Element = this.domElement.querySelector(`.useful_list`);
    resultContainer.innerHTML = '';

    const xml = `<View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name="ShowInWP" />
                                <Value Type="Boolean">1</Value>
                            </Eq>
                        </Where>
                        <OrderBy>
                            <FieldRef Name="Index" Ascending="TRUE" />
                        </OrderBy>
                    </Query>
                    <RowLimit>100</RowLimit>
               </View>`;

    const q: any = {
      ViewXml: xml,
    };

    web.get().then(w => {
      web.lists.getByTitle(usefulLinksList).getItemsByCAMLQuery(q).then((r: any[]) => {
        let html = '';
        for (let idx = 0; idx < r.length; idx++) {

          let result = r[idx];
          if (idx % 3 == 0) {
            html += '';
          }
          html += `
                        <div class='item'>
                           <div>
                              <img src='${result.Image.Url}'/>
                           </div>
                            <a  target='_blank'  href='${result.Link}'>${result.Title}</a>
                       </div>
                  `;
          if (idx % 3 == 2) {
            html += '</div>';
          }
        }
        //});
        resultContainer.insertAdjacentHTML('beforeend', html);
      })
        .catch(console.log);
    });
  }

  public Build(listName): string {

    let eventDay = parseInt(moment().add(-(parseInt(this.properties.daysBefore)), 'd').format('DD'));
    let eventMonth = parseInt(moment().add(-(parseInt(this.properties.daysBefore)), 'd').format('MM'));

    let nextDate = moment().add((parseInt(this.properties.daysAfter)), 'd');
    let nextEventMonth = parseInt(nextDate.format('MM'));
    let nextEventDay = parseInt(nextDate.format('DD'));


    let firstMonthDays = [];
    let secondMonthDays = [];

    //next date is on the same month
    if (nextEventMonth == eventMonth) {
      let _i1 = 0;
      for (let i1 = eventDay; i1 <= nextEventDay; i1++) {
        firstMonthDays[_i1] = i1;
        _i1 = _i1 + 1;
      }
    } else {

      //next date is on next month
      let _i = 0;
      let _j = 0;
      let tempDate = moment().add(-(parseInt(this.properties.daysBefore)), 'd');
      while (tempDate.isSame(nextDate) || tempDate.isBefore(nextDate)) {
        if (parseInt(tempDate.format('MM')) == eventMonth) {
          firstMonthDays[_i] = parseInt(tempDate.format('DD'));
          _i += 1;
        } else {
          secondMonthDays[_j] = parseInt(tempDate.format('DD'));
          _j += 1;
        }

        tempDate = tempDate.add(1, 'd');
      }
    }

    let _tempQuery = '';

    if (secondMonthDays.length == 0) {

      //build query for current month
      let _in = '';
      for (let i = 0; i < firstMonthDays.length; i++) {
        _in += '<Value Type="Number">' + firstMonthDays[i] + '</Value>';
      }
      _tempQuery =
        '<And>' +
        '<And>' +
        '<Eq>' +
        '<FieldRef Name="EventMonth" />' +
        '<Value Type="Number">' + eventMonth + '</Value>' +
        '</Eq>' +
        '<In>' +
        '<FieldRef Name="EventDay" />' +
        '<Values>' + _in + '</Values>' +
        '</In>' +
        '</And>' +
        '<Or>' +
        '<Or>' +
        '<IsNull>' +
        '<FieldRef Name="Expires" />' +
        '</IsNull>' +
        '<And>' +
        '<IsNull>' +
        '<FieldRef Name="EventType" />' +
        '</IsNull>' +
        '<Eq>' +
        '<FieldRef Name="EventType" />' +
        '<Value Type="Text">birthday</Value>' +
        '</Eq>' +
        '</And>' +
        '</Or>' +
        '<Gt>' +
        '<FieldRef Name="Expires" />' +
        '<Value Type="DateTime">' +
        '<Today />' +
        '</Value>' +
        '</Gt>' +
        '</Or>' +
        '</And>';

    } else {

      //build query for 2 months
      let _in_first = '';
      let _in_second = '';

      for (let i3 = 0; i3 < firstMonthDays.length; i3++) {
        _in_first += '<Value Type="Number">' + firstMonthDays[i3] + '</Value>';
      }
      for (let i4 = 0; i4 < secondMonthDays.length; i4++) {
        _in_second += '<Value Type="Number">' + secondMonthDays[i4] + '</Value>';
      }
      _tempQuery =
        '<Or>' +
        '<And>' +
        '<And>' +
        '<Eq>' +
        '<FieldRef Name="EventMonth" />' +
        '<Value Type="Number">' + eventMonth + '</Value>' +
        '</Eq>' +
        '<In>' +
        '<FieldRef Name="EventDay" />' +
        '<Values>' + _in_first + '</Values>' +
        '</In>' +
        '</And>' +
        '<Or>' +
        '<Or>' +
        '<IsNull>' +
        '<FieldRef Name="Expires" />' +
        '</IsNull>' +
        '<And>' +
        '<IsNull>' +
        '<FieldRef Name="EventType" />' +
        '</IsNull>' +
        '<Eq>' +
        '<FieldRef Name="EventType" />' +
        '<Value Type="Text">birthday</Value>' +
        '</Eq>' +
        '</And>' +
        '</Or>' +
        '<Gt>' +
        '<FieldRef Name="Expires" />' +
        '<Value Type="DateTime">' +
        '<Today />' +
        '</Value>' +
        '</Gt>' +
        '</Or>' +
        '</And>' +
        '<And>' +
        '<And>' +
        '<Eq>' +
        '<FieldRef Name="EventMonth" />' +
        '<Value Type="Number">' + nextEventMonth + '</Value>' +
        '</Eq>' +
        '<In>' +
        '<FieldRef Name="EventDay" />' +
        '<Values>' + _in_second + '</Values>' +
        '</In>' +
        '</And>' +
        '<Or>' +
        '<Or>' +
        '<IsNull>' +
        '<FieldRef Name="Expires" />' +
        '</IsNull>' +
        '<And>' +
        '<IsNull>' +
        '<FieldRef Name="EventType" />' +
        '</IsNull>' +
        '<Eq>' +
        '<FieldRef Name="EventType" />' +
        '<Value Type="Text">birthday</Value>' +
        '</Eq>' +
        '</And>' +
        '</Or>' +
        '<Gt>' +
        '<FieldRef Name="Expires" />' +
        '<Value Type="DateTime">' +
        '<Today />' +
        '</Value>' +
        '</Gt>' +
        '</Or>' +
        '</And>' +
        '</Or>';
    }

    //get items for currend date +- 3 days
    let tempQuery = '<View><ViewFields>' +
      '<FieldRef Name="ID"/><FieldRef Name="Title"/><FieldRef Name="EventType"/><FieldRef Name="BabyGender"/><FieldRef Name="EventAuthor"/><FieldRef Name="EventMonth"/><FieldRef Name="EventDay"/><FieldRef Name="Role"/><FieldRef Name="Expires"/>' +
      (listName == this._newEmployeesListName ? '<FieldRef Name="Picture"/>' : '') +
      '</ViewFields>' +
      '<Query>' +
      '<Where>' + _tempQuery +
      '</Where>' +
      '<OrderBy>' +
      '<FieldRef Name="EventMonth" Ascending="TRUE" />' +
      '<FieldRef Name="EventDay" Ascending="TRUE" />' +
      '</OrderBy>' +
      '</Query>' +
      '<RowLimit>10000</RowLimit></View>';


    return tempQuery;
  }

  // =======================================Functions of photo gallery webpart=====================================
  public static getAbsoluteDomainUrl(): string {
    if (window
      && 'location' in window
      && 'protocol' in window.location
      && 'host' in window.location) {
      return window.location.protocol + '//' + window.location.host;
    }
    return null;
  }

  private _getFiles(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/Web/GetFolderByServerRelativeUrl('${escape(this.properties.photosLibraryName)}')?$expand=Files`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getPropertiesFields() {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/Web/GetFolderByServerRelativeUrl('${escape(this.properties.photosLibraryName)}')/Files?$select=ListItemAllFields/Date,ListItemAllFields/ImageTitle,ListItemAllFields/Link&$expand=ListItemAllFields`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderListAsync(): void {
    if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getFiles()
        .then((response) => {
          this._renderList1(response.Files);
        });
      this._getPropertiesFields()
        .then((response) => {
          this._renderList2(response.value);
        });
    }
  }

  private _renderList1(items: ISPList[]): void {
    let html: string = '';
    let index: number = 0;
    items.forEach((item: ISPList) => {
      if (index == 0) {
        html += `
          <div class="image-item" style="border-left: 1px solid lightgray;">
              <img src="${LumenisAllWebPart.getAbsoluteDomainUrl()}${item.ServerRelativeUrl}">
          </div> <div class="divider"></div>`;

      } else {
        html += `
          <div class="image-item" style="border-left: 1px solid lightgray;"><img src="${LumenisAllWebPart.getAbsoluteDomainUrl()}${item.ServerRelativeUrl}"></div><div class="divider"></div>`;
      }
      index++;
    });
    const slidesShowContainer: Element = this.domElement.querySelector('#slidesShow-Container');
    slidesShowContainer.innerHTML = html;

    let btnNext = document.getElementById('prev');
    btnNext.addEventListener('click', (e: Event) => this.plusSlides(-1));
    let btnPrev = document.getElementById('next');
    btnPrev.addEventListener('click', (e: Event) => this.plusSlides(1));
  }

  private _renderList2(items: ISPList[]): void {
    let html: string = '';
    let html2: string = '';
    let num = -8.5;
    const mySlides: Element = this.domElement.querySelector('#mySlides');
    const slidesShowContainer: Element = this.domElement.querySelector('#FileProperties');
    const afterSlidesShowContainer: Element = this.domElement.querySelector('#after-slidesShow-Container');

    items.forEach((item: ISPList) => {
      num = num + 10;
      html += `
          <div >
          <img src="https://lumenis.sharepoint.com/sites/Portal/Site%20Assets/Homepage%20design/image%20gallery/icon_camera.jpg" alt="icon" style="float:right;">
          <label>${item.ListItemAllFields.Date}</label>
          <br><label>${item.ListItemAllFields.ImageTitle}</label></div>`;

      html2 += `<div ><a href="${item.ListItemAllFields.Link}"> קרא עוד ></a></div>`;


    });

    slidesShowContainer.innerHTML = html;
    afterSlidesShowContainer.innerHTML = html2;
  }

  private plusSlides(n: number) {
    this.showSlides(this.slideIndex += n);
  }

  private showSlides(n: number): void {
    var i: number;
    var slides = document.getElementsByClassName('mySlides');
    if (n > slides.length) {
      this.slideIndex = 1;
    }
    if (n < 1) {
      this.slideIndex = slides.length;
    }
    for (i = 0; i < slides.length; i++) {
      slides[i].setAttribute('style', 'display:none;');
    }
    slides[this.slideIndex - 1].setAttribute('style', 'display:block;');
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Main Banner',
              groupFields: [
                PropertyPaneTextField('mainBannerImg', {
                  label: 'Enter Link to image'
                }),
              ]
            },
            {
              groupName: 'Useful Links',
              groupFields: [
                PropertyPaneTextField('descriptionUsefulLinks', {
                  label: 'Description field'
                }),
                PropertyPaneTextField('wpTitle_UsefulLinks', {
                  label: 'Title'
                }),
                PropertyPaneTextField('listName_UsefulLinks', {
                  label: 'List name'
                }),
                PropertyPaneTextField('wpWebUrl_UsefulLinks', {
                  label: 'Web URL'
                })
              ]
            },
            {
              groupName: 'Quick Link 1',
              groupFields: [
                PropertyPaneTextField('Link1', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage1', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link1Text', {
                  label: 'Enter link description'
                })
              ]
            },
            {
              groupName: 'Quick Link 2',
              groupFields: [
                PropertyPaneTextField('Link2', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage2', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link2Text', {
                  label: 'Enter link description'
                })
              ]
            },
            {
              groupName: 'Quick Link 3',
              groupFields: [
                PropertyPaneTextField('Link3', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage3', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link3Text', {
                  label: 'Enter link description'
                })
              ]
            },
            {
              groupName: 'Quick Link 4',
              groupFields: [
                PropertyPaneTextField('Link4', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage4', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link4Text', {
                  label: 'Enter link description'
                })
              ]
            },
            {
              groupName: 'Quick Link 5',
              groupFields: [
                PropertyPaneTextField('Link5', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage5', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link5Text', {
                  label: 'Enter link description'
                })
              ]
            },
            {
              groupName: 'Quick Link 6',
              groupFields: [
                PropertyPaneTextField('Link6', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage6', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link6Text', {
                  label: 'Enter link description'
                })
              ]
            },
            {
              groupName: 'Quick Link 7',
              groupFields: [
                PropertyPaneTextField('Link7', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage7', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link7Text', {
                  label: 'Enter link description'
                })
              ]
            },
            {
              groupName: 'Quick Link 8',
              groupFields: [
                PropertyPaneTextField('Link8', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage8', {
                  label: 'Enter link to image'
                }),
                PropertyPaneTextField('Link8Text', {
                  label: 'Enter link description'
                })
              ]
            },

            // Greetings
            {
              groupName: 'Greetings',
              groupFields: [
                PropertyPaneTextField('descriptionGreetings', {
                  label: 'Description field'
                }),
                PropertyPaneTextField('webUrlGreetings', {
                  label: 'Web url'
                }),
                PropertyPaneTextField('listNameGreetings', {
                  label: 'List name'
                }),
                PropertyPaneTextField('wpTitleGreetings', {
                  label: 'Wp title'
                }),
                PropertyPaneTextField('sendYourGreetings', {
                  label: 'Send your greetings'
                }),
                PropertyPaneTextField('separator', {
                  label: 'Separator'
                }),
                PropertyPaneTextField('daysBefore', {
                  label: 'Days before'
                }),
                PropertyPaneCheckbox('displayDate', {
                  text: 'Display date'
                }),
                PropertyPaneTextField('daysAfter', {
                  label: 'Days after'
                })
              ]
            },
            // Photo gallery

            {
              groupFields: [
                PropertyPaneTextField('photoGalleryTitle', {
                  label: 'Enter photos gallery element title'
                }),
                PropertyPaneTextField('photosLibraryName', {
                  label: 'Enter photos library name'
                })
              ]
            },
            {
              groupName: 'Link 1',
              groupFields: [
                PropertyPaneTextField('linksImagesTitle', {
                  label: 'Enter links images element title'
                }),
                PropertyPaneTextField('Link1PhotoGallery', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('LinkImage1PhotoGallery', {
                  label: 'Enter link to image'
                }),
              ]
            },
            {
              groupName: 'Link 2',
              groupFields: [
                PropertyPaneTextField('Link2PhotoGallery', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('Link2TextPhotoGallery', {
                  label: 'Enter text to display'
                }),
              ]
            },
            {
              groupName: 'Link 3',
              groupFields: [
                PropertyPaneTextField('Link3PhotoGallery', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('Link3TextPhotoGallery', {
                  label: 'Enter text to display'
                }),
              ]
            },
            {
              groupName: 'Link 4',
              groupFields: [
                PropertyPaneTextField('Link4PhotoGallery', {
                  label: 'Enter link to page'
                }),
                PropertyPaneTextField('Link4TextPhotoGallery', {
                  label: 'Enter text to display'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
