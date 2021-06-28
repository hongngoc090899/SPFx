import { UrlQueryParameterCollection, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { cloneDeepWith, escape } from '@microsoft/sp-lodash-subset';

import styles from './TestFormWebPart.module.scss';
import * as strings from 'TestFormWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import pnp, { sp, Item, ItemAddResult, ItemUpdateResult, Web } from 'sp-pnp-js';

import * as $ from 'jquery';
require('bootstrap');
require('./css/jquery-ui.css');
let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
require('appjs');
require('sppeoplepicker');
require('jqueryui');

var queryParms = new UrlQueryParameterCollection(window.location.href);
var SpId = queryParms.getValue("ID");

export interface ITestFormWebPartProps {
  description: string;
}

export default class TestFormWebPart extends BaseClientSideWebPart<ITestFormWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div id="container" class="container">
      <div class="panel">
          <div class="panel-body">
              <div class="row">
                  <div class="col-lg-3 control-padding">
                      <label>Activity</label>
                      <input type='textbox' name='txtActivity' id='txtActivity' class="form-control" value=""
                          placeholder="">
                  </div>
                  <div class="col-lg-3 control-padding">
                      <label>Activity Performed By</label>

                      <div id="ppDefault"></div>
                  </div>
              </div>

              <div class="row">
                <div class="col-lg-3 control-padding">
                  <label>Activity Date</label>
                  <div class="input-group date" data-provide="datepicker">
                      <input type="text" class="form-control" id="txtDate" name="txtDate">
                  </div>
                </div>
              </div>

              <div class="row">
                  <div class="col-lg-4 control-padding">
                      <label>Category</label>
                      <select name="ddlCategory" id="ddlCategory" class="form-control">

                      </select>
                  </div>
              </div>

              <div class="row">
                <div class="col-lg-10">
                <label>Note</label>
                <div id="ddlSubCategory">

                </div>
                </div>
              </div>

              <div class="row" style="margin-top:1rem;">
                  <div class="col col-lg-12">
                      <button type="button" class="btn btn-primary buttons" id="btnSubmit">Save</button>
                      <button type="button" class="btn btn-default buttons" id="btnCancel">Cancel</button>
                  </div>
              </div>
          </div>
          </div>
          <div class="panel">
          <ul>
            <li>Form dùng để CRUD dữ liệu từ list</li>
            <li>Trường Person lấy dữ liệu account user theo tên</li>
            <li>Trường Time để chọn ngày</li>
            <li>Trường Category để lấy dữ liệu từ column 'Gift name' trong list 'Gift'</li>
            <li>Trường Note để lấy dữ liệu từ list 'Power app form' column 'Multiple lines of text' và lọc theo column 'People User' đã nhập ở trường Person</li>
          </ul>
        </div>
      </div>`;

    (<any>$("#txtDate")).datepicker(
      {
        changeMonth: true,
        changeYear: true,
        dateFormat: "mm/dd/yy"
      }
    );
    (<any>$('#ppDefault')).spPeoplePicker({
      minSearchTriggerLength: 2,
      maximumEntitySuggestions: 10,
      principalType: 1,
      principalSource: 15,
      searchPrefix: '',
      searchSuffix: '',
      displayResultCount: 6,
      maxSelectedUsers: 1
    });
    this.AddEventListeners();
    this.getCategoryData();
    // this.getSubCategoryData();
  }

  private AddEventListeners(): any {
    document.getElementById('btnSubmit').addEventListener('click', () => this.SubmitData());
    document.getElementById('btnCancel').addEventListener('click', () => this.CancelForm());
    document.getElementById('ppDefault').addEventListener('change', () => this.getSubCategoryData());
  }

  private SubmitData(){
    var userinfo = (<any>$('#ppDefault')).spPeoplePicker('get');
    var userDetails = this.GetUserId(userinfo[0].email.toString());
    var userId = userDetails.d.Id;

    var test = {
      "Title": "Test",
      "Activity": $("#txtActivity").val().toString(),
      "Activity_Date": $("#txtDate").val().toString(),
      "Activity_ById" : userId,
      "Category" : $("#ddlCategory").html(),
      "SubCategory": $("#ddlSubCategory").val().toString(),
    }

    console.log(test);

    pnp.sp.web.lists.getByTitle('TestCRUDSPFx').items.add({
      Title: $("#txtActivity").val().toString(),
      Person : userId,
      Time: $("#txtDate").val().toString(),
      Category: $("#ddlCategory").val().toString(),
      Note: $("#ddlSubCategory").val().toString(),
    });
  }

  private GetUserId(userName) {
    var siteUrl = this.context.pageContext.web.absoluteUrl;
    var call = $.ajax({
      url: siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" + userName + "%27",
      method: "GET",
      headers: { "Accept": "application/json; odata=verbose" },
      async: false,
      dataType: 'json'
    }).responseJSON;
    return call;
  }


  private CancelForm() {
    window.location.href = this.GetQueryStringByParameter("Source");
  }

  private GetQueryStringByParameter(name) {
    name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
    var regex = new RegExp("[\\?&]" + name + "=([^&#]*)"),
      results = regex.exec(location.search);
    return results == null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
  }

  private _getCategoryData(): any {
    return pnp.sp.web.lists.getByTitle("Gift").items.select('Gift_x0020_Name').get().then((response) => {
      return response;
    });
  }

  private getCategoryData(): any {
    this._getCategoryData()
      .then((response) => {
        this._renderCategoryList(response);
      });
  }

  private _renderCategoryList(items: any): void {

    let html: string = '';
    html += `<option value="Select Category" selected>Select Category</option>`;
    items.forEach((item: any) => {
      html += `
       <option value="${item.Gift_x0020_Name}">${item.Gift_x0020_Name}</option>`;
    });
    const listContainer1: Element = this.domElement.querySelector('#ddlCategory');
    listContainer1.innerHTML = html;
  }

  private _getSubCategoryData(): any {

    return pnp.sp.web.lists.getByTitle("TestPowerApp").items.get().then((response) => {
      var userinfo = (<any>$('#ppDefault')).spPeoplePicker('get');
      var userDetails = this.GetUserId(userinfo[0].email.toString());
      var userId = userDetails.d.Id;

      var listNote = [];

      response.forEach(items => {
        if(userId === items.Person_x0020_UserId) {
          var note = {
            "note" : items.Multiplelinesoftext
          }
          listNote.push(note)
        }
      });
      return listNote;
    }).catch((err)=>{
      var a = err;
    });
  }

  private getSubCategoryData(): any {
    this._getSubCategoryData()
      .then((response) => {
        this._renderSubCategoryList(response);
      });
  }

  private _renderSubCategoryList(items: any): void {

    let html: string = '';
    // html += `<option value="Select Category" selected>Select Category</option>`;
    items.forEach((item: any) => {
      // html += `
      //  <option value="${item.GiftName}">${item.GiftName}</option>`;
      html += item.note;
    });
    const listContainer1: Element = this.domElement.querySelector('#ddlSubCategory');
    listContainer1.innerHTML = html;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

