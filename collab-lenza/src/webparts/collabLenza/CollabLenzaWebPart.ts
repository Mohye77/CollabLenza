import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  escape
} from '@microsoft/sp-lodash-subset';
import {
  Version,
  Environment,
  EnvironmentType,
  Log
} from '@microsoft/sp-core-library';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import Web from 'sp-pnp-js';
import styles from './CollabLenzaWebPart.module.scss';
import * as strings from 'CollabLenzaWebPartStrings';

export interface ICollabLenzaWebPartProps {
  description: string;
  drawId: string;
  listTitle: string;
}
import {
  IListItem
} from './IListItem';

export default class CollabLenzaWebPart extends BaseClientSideWebPart < ICollabLenzaWebPartProps > {
  private _canvas: HTMLCanvasElement;
  private _context2D: CanvasRenderingContext2D;
  private _userDrawings = [];
  private _allUserDrawings = [];
  private _pencilColor = '#000';
  private _penShape = 'line';
  private _pencilSize = 2;
  private _penBoundary = 'butt';
  private _pos1 = {
    'posX': 0,
    'posY': 0
  };
  private _pos2 = {
    'posX': 0,
    'posY': 0
  };
  private _layerId = 0;
  private _currentUser = '';
  private _listItemEntityTypeName = '';
  private _spItemID = -1;
  private _etag: string;


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('description', {
              label: strings.DescriptionFieldLabel
            }),
            PropertyPaneTextField('drawId', {
              label: strings.DrawIdFieldLabel
            })
          ]
        }]
      }]
    };
  }

  private _generateGUID = (): string => {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = Math.random() * 16 | 0,
        v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }

  private _matches = (element, selector): boolean => {
    return (element.matches || element.matchesSelector || element.msMatchesSelector || element.mozMatchesSelector || element.webkitMatchesSelector || element.oMatchesSelector).call(element, selector);
  };

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.collabLenza }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
              <div class="${ styles.column }">
                <span class="${ styles.title }">${escape(this.title)}</span>
                <p class="${ styles.description }">${escape(this.properties.description)}</p>
                <canvas id="allUsersCanvas" class="${ styles.canvas }">
                </canvas>
                <div class="${ styles.palette }">
                    <div class="${ styles.pencil } black selected" data-color="#000"></div>
                    <div class="${ styles.pencil } white" data-color="#fff"></div>
                    <div class="${ styles.pencil } blue" data-color="#00f"></div>
                    <div class="${ styles.pencil } green" data-color="#0f0"></div>
                    <div class="${ styles.pencil } yellow" data-color="#ff0"></div>
                    <div class="${ styles.pencil } orange" data-color="#fa0"></div>
                    <div class="${ styles.pencil } red" data-color="#f00"></div>
                    <div class="${ styles.pencil } purple" data-color="#f0f"></div>
                </div>
                <div class="${ styles.tools }">
                    <div class="${ styles.tool } selected" data-tool="line">
                      <img src="/images/line.png" alt="Line" /></div>
                    <div class="${ styles.tool }" data-tool="circle">
                        <img src="/images/circle.png" alt="Circle" /></div>
                    <div class="${ styles.tool }" data-tool="rectangle">
                        <img src="/images/rectangle.png" alt="Rectangle" /></div>
                    <div class="${ styles.size }">
                        <input id="pencilSize" type="range" min="0" max="10" step="1" value="2" />
                        <span id="pencilSizeValue">2</span>
                    </div>
                    <div class="${ styles.actions }">
                        <button class="undo">
                            <img src="/images/undo.png" alt="Undo" /></button>
                        <button class="redo">
                            <img src="/images/redo.png" alt="Redo" /></button>
                        <button class="sync">
                            <img src="/images/sync.png" alt="Save & sync" /></button>
                        <button class="save">
                            <img src="/images/save.png" alt="Get as image" /></button>
                    </div>
                    <div class="${ styles.image }">
                        <img id="image" />
                    </div>
                </div>
              </div>
          </div>
        </div>
    </div>`;

    this._initialize();
  }

  private _initialize = () => {
    this._currentUser = this.context.pageContext.user.loginName;
    this._canvas = < HTMLCanvasElement > document.getElementById('allUsersCanvas');
    this._context2D = this._canvas.getContext('2d');

    //#region Initialize Context2D
    this._context2D.lineWidth = this._pencilSize;
    this._context2D.lineCap = this._penBoundary;
    //#endregion Initialize Context2D

    //#region Initialize Events
    this._canvas.addEventListener('mousedown', this._getCursorPosition, false);
    this._canvas.addEventListener('mouseup', this._draw, false);

    let pencilColors = document.querySelectorAll('.pencil');
    for (var i = 0; i < pencilColors.length; i++) {
      pencilColors[i].addEventListener('click', this._selectColor, false);
    }

    let penShapes = document.querySelectorAll('.' + styles.tool);
    for (var i = 0; i < penShapes.length; i++) {
      penShapes[i].addEventListener('click', this._selectShape, false);
    }

    let pencilSizeInput = document.querySelector('input#pencilSize');
    pencilSizeInput.addEventListener('change', this._selectSize, false);

    let undoBtn = document.querySelector('button.undo');
    undoBtn.addEventListener('click', this._undo, false);

    let redoBtn = document.querySelector('button.redo');
    redoBtn.addEventListener('click', this._redo, false);

    let self = this;
    let syncBtn = document.querySelector('button.sync');
    syncBtn.addEventListener('click', function (event) {
      self._sync(false);
    }, false);

    let saveBtn = document.querySelector('button.save');
    saveBtn.addEventListener('click', this._save, false);
    //#endregion Events

    this._sync(false);
  }

  //#region DataAccess
  private _sync(updateCurrentUserDrawings: boolean): void {
    if (updateCurrentUserDrawings === true) {
      if (this._spItemID === -1) {
        this._getListItemEntityTypeName()
          .then((listItemEntityTypeName: string): Promise < SPHttpClientResponse > => {
            const body: string = JSON.stringify({
              '__metadata': {
                'type': listItemEntityTypeName
              },
              'WebPartId': this.properties.drawId,
              'UserId': this._currentUser,
              'Layers': JSON.stringify(this._userDrawings)
            });
            return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listTitle}')/items`,
              SPHttpClient.configurations.v1, {
                headers: {
                  'Accept': 'application/json;odata=nometadata',
                  'Content-type': 'application/json;odata=verbose',
                  'odata-version': ''
                },
                body: body
              });
          })
          .then((response: SPHttpClientResponse): Promise < IListItem > => {
            return response.json();
          })
          .then((item: IListItem): void => {
            console.log(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
            this._spItemID = item.Id;
            this._computeDrawings();
          }, (error: any): void => {
            console.log('Error while creating the item: ' + error);
          });
      } else {
        this._getListItemEntityTypeName()
          .then((listItemType: string): Promise < SPHttpClientResponse > => {
            this._listItemEntityTypeName = listItemType;

            console.log(`Loading information about item ID: ${this._spItemID}...`);
            return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listTitle}')/items(${this._spItemID})?$select=Id`,
              SPHttpClient.configurations.v1, {
                headers: {
                  'Accept': 'application/json;odata=nometadata',
                  'odata-version': ''
                }
              });
          })
          .then((response: SPHttpClientResponse): Promise < IListItem > => {
            this._etag = response.headers.get('ETag');
            return response.json();
          })
          .then((item: IListItem): Promise < SPHttpClientResponse > => {
            console.log(`Updating item with ID: ${this._spItemID}...`);
            const body: string = JSON.stringify({
              '__metadata': {
                'type': this._listItemEntityTypeName
              },
              'Layers': JSON.stringify(this._userDrawings)
            });
            return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listTitle}')/items(${item.Id})`,
              SPHttpClient.configurations.v1, {
                headers: {
                  'Accept': 'application/json;odata=nometadata',
                  'Content-type': 'application/json;odata=verbose',
                  'odata-version': '',
                  'IF-MATCH': this._etag,
                  'X-HTTP-Method': 'MERGE'
                },
                body: body
              });
          })
          .then((response: SPHttpClientResponse): void => {
            console.log(`Item with ID: ${this._spItemID} successfully updated`);
          }, (error: any): void => {
            console.log(`Error updating item: ${error}`);
          });
      }


    }
    //on récupère les dessins
  }

  private _computeDrawings = () => {
    this._getDrawings(this.properties.listTitle).then((response) => {
      this._renderDrawings(response);
    }).catch((err) => {
      Log.error('_initializeDrawings', err);
      this.context.statusRenderer.renderError(this.domElement, err);
    });
  }

  private _getDrawings(listName: string): Promise < IListItem[] > {
    const queryString: string = `$filter=(WebPartID eq '${this.properties.drawId}')$select=ID,UserID,WebPartID,Layers`;

    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items?${queryString}`,
        SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 404) {
          Log.error('_getDrawings', new Error('List not found.'));
          return [];
        } else {
          return response.json();
        }
      });
  }

  private _renderDrawings = (items: IListItem[]) => {
    items.forEach((item: IListItem) => {
      let userID: string = item["UserID"];
      let layers: string = item["Layers"];
      let webPartID: string = item["WebPartID"];
      let spItemID: string = item["ID"];

      if (this._currentUser == userID) {
        this._spItemID = parseInt(spItemID);
        this._userDrawings = JSON.parse(layers);
      } else {
        if (this._allUserDrawings.length == 0) {
          this._allUserDrawings = JSON.parse(layers);
        } else {
          this._allUserDrawings = this._allUserDrawings.concat(JSON.parse(layers));
        }
      }

    });
    this._redraw(true);
  }

  private _getListItemEntityTypeName(): Promise < string > {
    return new Promise < string > ((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this._listItemEntityTypeName) {
        resolve(this._listItemEntityTypeName);
        return;
      }

      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listTitle}')?$select=ListItemEntityTypeFullName`,
          SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse): Promise < {
          ListItemEntityTypeFullName: string
        } > => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: {
          ListItemEntityTypeFullName: string
        }): void => {
          this._listItemEntityTypeName = response.ListItemEntityTypeFullName;
          resolve(this._listItemEntityTypeName);
        });
    });
  }
  //#endregion DataAccess

  private _getCursorPosition = (event) => {
    let rect = this._canvas.getBoundingClientRect();
    let posX = event.clientX - rect.left;

    let posY = event.clientY - rect.top;

    this._pos1 = {
      'posX': posX,
      'posY': posY
    };
  }

  private _selectColor = (event) => {
    let selectedPencilColor = document.querySelector('.' + styles.pencil + '.selected');
    if (selectedPencilColor) {
      selectedPencilColor.classList.remove('selected');
    }

    if (this._matches(event.target, '.' + styles.pencil)) {
      event.target.classList.add('selected');
      this._pencilColor = event.target.getAttribute('data-color');
    } else {
      event.target.parentNode.classList.add('selected');
      this._pencilColor = event.target.parentNode.getAttribute('data-color');
    }
  }

  private _selectShape = (event) => {
    let selectedpenShape = document.querySelector('.' + styles.tool + '.selected');
    if (selectedpenShape) {
      selectedpenShape.classList.remove('selected');
    }

    if (this._matches(event.target, '.' + styles.tool)) {
      event.target.classList.add('selected');
      this._penShape = event.target.getAttribute('data-tool');
    } else {
      event.target.parentNode.classList.add('selected');
      this._penShape = event.target.parentNode.getAttribute('data-tool');
    }
  }

  private _selectSize = (event) => {
    this._pencilSize = event.target.parentNode.value;
    let pencilSizeSpan = document.querySelector('span#pencilSizeValue');
    pencilSizeSpan.innerHTML = this._pencilSize.toString();
  }

  private _undo = (event) => {
    if (this._layerId > -1) {
      let drawing = this._userDrawings[this._layerId];
      if (drawing != undefined && drawing != null) {
        drawing.visible = false;
        this._redraw(false);
        if (this._layerId > 0) {
          this._layerId--;
        }
      }
    }
  }

  private _redo = (event) => {
    if (this._layerId < this._userDrawings.length - 1) {
      this._layerId++;
      let drawing = this._userDrawings[this._layerId];
      if (drawing != undefined && drawing != null) {
        drawing.visible = true;
        this._redraw(false);
      }
    }
  }

  private _save = (event) => {
    var imageUrl = this._canvas.toDataURL("image/png");
    var image = document.querySelector("img#image");
    image.setAttribute('src', imageUrl);
    image.setAttribute('style', 'display:block;');
  }

  private _draw = (event) => {
    let layerName = this._generateGUID();
    let rect = this._canvas.getBoundingClientRect();
    let posX = event.clientX - rect.left;
    let posY = event.clientY - rect.top;

    this._pos2 = {
      'posX': posX,
      'posY': posY
    };

    let drawing = {
      'layerName': layerName,
      'pos1': this._pos1,
      'pos2': this._pos2,
      'penShape': this._penShape,
      'pencilColor': this._pencilColor,
      'pencilSize': this._pencilSize,
      'penBoundary': this._penBoundary,
      'date': new Date(),
      'visible': true
    };
    /*
    'layerName' : ID de l'action
    'pos1' : Position initiale (au clic)
    'pos2' : Position finale (au relâchement)
    'penShape' : La forme choisie (ligne, cercle, rectangle)
    'penColor' : La couleur choisie (en hexadécimal)
    'penSize' : La largeur choisie
    'penBoundary' : La forme de l'extrémité
    'date' : Le moment de l'action
    'visible' : L'action a-t'elle été annulé?
    */
    this._userDrawings.push(drawing);
    this._layerId = this._userDrawings.length - 1;
    this._drawFromLayer(drawing);
  }

  private _drawFromLayer = (drawing) => {
    this._context2D.strokeStyle = drawing.pencilColor;
    this._context2D.lineWidth = drawing.pencilSize;
    this._context2D.lineCap = drawing.penBoundary;

    this._context2D.beginPath();
    if (drawing.penShape == 'line') {
      this._drawLine(drawing);
    } else if (drawing.penShape == 'circle') {
      this._drawCircle(drawing);
    } else if (drawing.penShape == 'rectangle') {
      this._drawRectangle(drawing);
    }
    this._context2D.closePath();
  }

  private _drawLine = (drawing) => {
    this._context2D.moveTo(drawing.pos1.posX, drawing.pos1.posY);
    this._context2D.lineTo(drawing.pos2.posX, drawing.pos2.posY);
    this._context2D.stroke();
  }

  private _drawRectangle = (drawing) => {
    let pos = {
      'posX': 0,
      'posY': 0
    };
    let width = 0;
    let height = 0;
    if (drawing.pos1.posX > drawing.pos2.posX) {
      pos.posX = drawing.pos2.posX;
      width = drawing.pos1.posX - drawing.pos2.posX;
    } else {
      pos.posX = drawing.pos1.posX;
      width = drawing.pos2.posX - drawing.pos1.posX;
    }
    if (drawing.pos1.posY > drawing.pos2.posY) {
      pos.posY = drawing.pos2.posY;
      height = drawing.pos1.posY - drawing.pos2.posY;
    } else {
      pos.posY = drawing.pos1.posY;
      height = drawing.pos2.posY - drawing.pos1.posY;
    }
    this._context2D.strokeRect(pos.posX, pos.posY, width, height);
  }

  private _drawCircle = (drawing) => {
    let pos = {
      'posX': 0,
      'posY': 0
    };
    var radius = 0;
    if (drawing.pos1.posX > drawing.pos2.posX) {
      radius = (drawing.pos1.posX - drawing.pos2.posX) / 2;
      pos.posX = drawing.pos2.posX + radius;
    } else {
      radius = drawing.pos2.posX - drawing.pos1.posX;
      pos.posX = drawing.pos1.posX + radius;
    }
    if (drawing.pos1.posY > drawing.pos2.posY) {
      pos.posY = drawing.pos2.posY + radius;
    } else {
      pos.posY = drawing.pos1.posY + radius;
    }
    this._context2D.arc(pos.posX, pos.posY, radius, 0, Math.PI * 2, true);
    this._context2D.stroke();
  }

  private _redraw = (fromSync: boolean) => {
    this._context2D.clearRect(0, 0, this._canvas.width, this._canvas.height);
    if (fromSync === true) {
      this._layerId = this._userDrawings.length - 1;
    }
    let allDrawings = this._allUserDrawings.concat(this._userDrawings);
    allDrawings = allDrawings.sort(function (drawing1, drawing2) {
      if (drawing1.date < drawing2.date)
        return -1;
      if (drawing1.date > drawing2.date)
        return 1;
      return 0;
    });
    for (var i = 0; i < allDrawings.length; i++) {
      if (allDrawings[i].visible === true) {
        this._drawFromLayer(allDrawings[i]);
      }
    }
  }
}
