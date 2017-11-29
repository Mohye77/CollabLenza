import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  Version,
  Environment,
  EnvironmentType,
  Log
} from '@microsoft/sp-core-library';
import {
  default as pnp,
  ItemAddResult
} from "sp-pnp-js";
import {
  IListItem
} from './IListItem';
import styles from './CollabLenzaWebPart.module.scss';
import * as strings from 'CollabLenzaWebPartStrings';
export interface ICollabLenzaWebPartProps {
  description: string;
  drawId: string;
  listTitle: string;
}

export default class CollabLenzaWebPart extends BaseClientSideWebPart < ICollabLenzaWebPartProps > {
  private _canvas: HTMLCanvasElement;
  private _context2D: CanvasRenderingContext2D;
  private _userDrawings = [];
  private _allUserDrawings = [];
  private _pencilColor = '#000';
  private _penShape = 'line';
  private _pencilSize = 2;
  private _penBoundary = 'butt';
  private _initialPosition = {
    'posX': 0,
    'posY': 0
  };
  private _finalPosition = {
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
                <span class="${ styles.title }">${this.title}</span>
                <p class="${ styles.description }">${this.properties.description}</p>
                <canvas id="allUsersCanvas" class="${ styles.canvas }">
                </canvas>
                <div class="${ styles.palette }">
                    <div class="${ styles.pencil } ${ styles.black } selected" data-color="#000"></div>
                    <div class="${ styles.pencil } ${ styles.white }" data-color="#fff"></div>
                    <div class="${ styles.pencil } ${ styles.blue }" data-color="#00f"></div>
                    <div class="${ styles.pencil } ${ styles.green }" data-color="#0f0"></div>
                    <div class="${ styles.pencil } ${ styles.yellow }" data-color="#ff0"></div>
                    <div class="${ styles.pencil } ${ styles.orange }" data-color="#fa0"></div>
                    <div class="${ styles.pencil } ${ styles.red }" data-color="#f00"></div>
                    <div class="${ styles.pencil } ${ styles.purple }" data-color="#f0f"></div>
                </div>
                <div class="${ styles.tools }">
                    <div class="${ styles.tool } selected" data-tool="line">
                      <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAFBhaW50Lk5FVCB2My41Ljg3O4BdAAAAnUlEQVRoQ+3ZsQ2DQBBE0SuFEhxSgktw6JASCOnEJRI6tHcQJ+EKPHP6T1qd2OxDdjQAcHM7z2j3mnfN63gK1SM+NZsWiYaIeNYogIh/I8IFES6IcLHW9IhFi0R6+z1CXyUSES6IcEGECyJcEOFCVzXDROjG46FFomuErm8iEeGCCBdDRew1sxaJrhGxP1yIcKKQ+IhuOk8AAH619gUplmTVdSU5sAAAAABJRU5ErkJggg==" alt="Line" /></div>
                    <div class="${ styles.tool }" data-tool="circle">
                        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAFBhaW50Lk5FVCB2My41Ljg3O4BdAAACYElEQVRoQ+2ZK0wEQRBETyKRSCQSiUQikUgkEonEIZFIJBKJRCKRSCQSiYR6l1SyuSzLzn+5TCWV++R2Z7q7+jN7q46Ojo5U7IjHI9wVF48z8V58F78n+Ck+ihfiYgxjI9fih7i54ecRfonD3/D5TtwXmwGP4l1v6lW8FA/FKbDpc/FJHBp0IyLJamAxpOFN4Gn0HwOMQo6+F86oEh0WYTEWJRonYg4ciW8i90Wmf0U1CRjhRC7hOfLNkUZqsVGeBHJyJNB2yWpzKzri2WVmT2FMjYT0esgtm9OoRMU89AuGCqAYJGNPdO3PldhzgdO8dnLyW6+EugWyrE943fCKlsMJDBVxwBcxoPtyAxpeSzgqvEbBlQODWoJmyT7oYVGwrGpVqilE7wVtciHjwhLgAfN0/SkAlFoubJ0fhvPkav0pAE70LM0oAzjzsB9eg8BZYysM2ZqIcP7mwlYdfROcHqMMoZNzIdPnEvAgsh8cHATGE48GS3jS4QNd1KhE6eXi4NqdGTRB9kFTjIITrHXC+zyEvKLg7o7EeN8KPmAlnYc8OEZPnolA1qyPMUlw9SIqtYdHCo6fYCKvZJAj3OxFrPHgweBRapZoGOSHyx83rwEnOEqIPhmOAYn5TJAlzBMgqd3DipR+j/awVPIzrNqI4JE9BIwIXoiGmavrk3vOieJGGMjM1QS5IbWUIoB8nIM4KXieSgGlePj/BobhxbklmkgiIzc7yPuoWSoH8N5wMxDvUrIZcTaJfMZ+X7qAzAZ/AbB5V7Y5ZGqoKqNQIA9OmGMRQU5F/vfo6Ojo+C9YrX4A6I3f4h7XMHIAAAAASUVORK5CYII=" alt="Circle" /></div>
                    <div class="${ styles.tool }" data-tool="rectangle">
                        <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAMAAAAp4XiDAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAMAUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALMw9IgAAAEAdFJOU////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wBT9wclAAAACXBIWXMAAA7DAAAOwwHHb6hkAAAAGXRFWHRTb2Z0d2FyZQBQYWludC5ORVQgdjMuNS44NzuAXQAAAEFJREFUSEvti7ENADAMg/L/0+lgtkqNs1ZmYYHqNR8uZUGLLGgRfpFFRlkmssgoy0QWGWWZuBYHWmRBK23IsqP7ACZTm9W3nJgNAAAAAElFTkSuQmCC" alt="Rectangle" /></div>
                    <div class="${ styles.size }">
                        <input id="pencilSize" type="range" min="1" max="10" step="1" value="2" />
                        <span id="pencilSizeValue">2</span>
                    </div>
                    <div class="${ styles.actions }">
                        <button class="undo">
                            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAFBhaW50Lk5FVCB2My41Ljg3O4BdAAACT0lEQVRoQ+3YTYhOURzH8WeQslDeElYsLJRkYTFDYcFeQg1RdtJslMUsZGGyYGupbJmatbwslJcQU7bICoUiUzNDxst8f3TquP09957z3Oc6t86vPs08T89z7vk/59x7z7mdnJycnJychLIASzGEwzjh0esdWAZ9rrFswcCff7tmIbZiDE8xh19d/MAkzmMbFqGvuY5h/KsYdWAnbmEaVqfLzOI29qBvBd3BDI79fvV3NmAcZb9+VT8xgY2oPSpEB/kCzW/Na43OPrxGsTN1eIeD0HStLa4Q0dTRyJyCCvMPXje1P4oq52el+IXIN3wvvNcvOtYZ1DIyxUKapvPvOHrO/y5EpjCInhJbyGdcw0nshe4x+qub4lV8gq5S1nctT7AY0Qkt5CPOYhW6ZQV0Mr+H1Y5lBNEJKeQZNiPkSqPPP4TVXtFzrERUqhaiaRL7i2l07sFqtyj6xA8ZES01jiBmMbgeL2G167uLqISeI7ppqpiY7IfVpu8r1iI4oYWIRuYoQkdGN77HsNp0NIUPIDgxhYgWmocQGp1nVnu+c0g+uorpR7AKcLRCTj66vL6BVYCjWZJ8tLF6AasApxWFLMErWAU4rShkHcqWLTeRfHajbOt8BcnnIqzO+04j6SzHB1idd/QYaReSzgVYnfdpm6ALQrLROktLG6vzvstIOjdgddynhxHJT6s1uA+rAEdPI3va7jaV1XgEqwiNxna0JtprPECxkEuo7WFdU9HI+MXoCYq2w62MRkbT7C026Y02R4+UtD/JycnJyclpIJ3OPC5vDxQOv1k5AAAAAElFTkSuQmCC" alt="Undo" /></button>
                        <button class="redo">
                            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAFBhaW50Lk5FVCB2My41Ljg3O4BdAAACUElEQVRoQ+3YO2hUQRTG8fWBYCEYFTFWsbAIiFhYRAW1ML1IIkRRsBOxESxSiEWChbaWgq0GUotJioAPNGjA1gQrFVRQFMwDk2j+38qF4XI2c+/szTor88GPZJPds3OY+5i5tZSUlJSUlMiyEdtxDOdw2aHXR7ANel902YzDuIVprODPGpbwGsM4hE3wZQMO/v21+qiBU5jAPKxB+/zEOI5D9ayoiQE8qr+qOPsxit+wBliWZmkE+5DPRcxhsv6qougw6McnWANq1nuchmZB55HOqwXof5U1ouKDyAqvF9W/Bs2EDr3s75U0opm4gV9wv3S9LCP/XZU0cgk6jt3CrdZ0Iz34Aat4KzXVyBa8glXYoqvYNzyAboK90L1CP6/gIb7D+qxPU41chVXU8hm6GOzAWtmFm/gKq04jwY3sxFtYRfNe4ACKRldAvf8NrHqW4EZ0glsF857CNwuNohkvelMNbuQJrIKuWXShbHSzO48yS5ugRjqxCKug6wxCoibcm10RQY30wTflUyiycnWjmbiAkEVmUCNDsIq5dHyXzVloAWjV8wlqRCtbq1hGgylzlfpnUfdWA5kP0OU5+vgamUGjjVBU8TXyDlsRfcZgNZDRcmQvos99WA1ktKQ/iehzHVYDrjuIPifge6zzBR2IOjqRiyyzbyP63IM1eJeWGqHrrZZFh1eRhw2PEXW0zdVTRGvwmWfYg+hzFI1m5SV2oy2iLeld5Jt4Du1Z2iraxrpPUtRE28xEPt34CB1ObTcT+Wj/oUc5KSkpKSkp/1FqtVUtrA8UtrOjFgAAAABJRU5ErkJggg==" alt="Redo" /></button>
                        <button class="sync">
                            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAACxIAAAsSAdLdfvwAAAAZdEVYdFNvZnR3YXJlAFBhaW50Lk5FVCB2My41Ljg3O4BdAAAFCElEQVRoQ+2Yb2xTVRjGzy2Nm13LWNt1Wzfae6djRg3DIGhQGSoiGDYTjUYJECOmUSQRCAKbjjMhAWNQo4agicHEEP3ARxNjYlQSvvj3k/MLGiIxURPjHEaYDOH6nNu35fbcc3tv2zsY2if55XbnvOd936e35/4Zq6uuS6Tu7u5lYMQH62hJ4DIMY569Fv7O4dhI0/6EBbuB6YPPaEngQuODUq3zGDvQ09MTphBvYdG0GUEzzbqudylIUYglhRGBMPM6jv7ODAKnzQjW7AUTMmjwUwqx5GJEYJnxdWYQPJ1GXpVyFPiCQiyVMSIQZjZRqLsQONONCHZTqLtEkLTIjboRBXUjrhJB0iI36kYU1I24SgRJi9z4Xxt5BVxQcGUZQYNprOuTia2JDTDOXgJNFHdpjfT392tBoI1ot7BRZoIf2S62Grnvl2rJBGrkDPihVvR5+onkiuSvZMQi8nhkMnNDRlWzQKBGqiZzfcaMD8TNhqcbTJyBogE7oeGQmViZMI1uQ5WjaiOnwR+1ovfqE7Mfnv03G2EXVM2raNjYcG5u31xx9u39VG1kGzBqoWl90wAMnFQ1W47QUOjDzoWdLyKHvZ+qjTxG09WJs7vQ1F9ykx6cxLrVYjnqi3d2q5dsb1YcazeCz/vAQb+0395+BGdiStGoqT2nTUY2RMaia6Nf2cbPwcB+EKWSRSOpZSlTG9bM1uWt79CUu7DAy8hP0pwruCKZ4S3hkuYFuNyaLYMtpt6jW3HpRWlrPLw1/DMMLKBSRSEml7w3aWIuv35Y+xOf4zStFhZ5GTkCjvoh9mjMsSfQxFTHko5v7HHYA1/C2HHjGmMjlSlRdF30PTkPjBykabWQOJg9wlkSBU9LDZzF+K0U4VupO1Oj+Bna8wimkCtLIU6haa8z0gSiXmhD2pBUWHyLuyhNRUK+XOK+RGmufL49FOIUFnkZOQaOe4Gb3Rmp8DgKW89SlQr5csa1hjnr2Vn2fIIxCnEKi2re7KIorlSlRTl7m1JULOS0rlrNDzaX5hzFjdVt02OBl5EnwOZytN/RvlcqKKh8n5GQ0zLS1t8m5xRfkHrPYUHtm52z2xwFR9lymq1YqG8Z6by5U84pGKCwUmGB1xlZD54qR2JVYr9cMD4Yf02KW0MpPYXYckasu79DWFDzHunq63IUxKO6HPc9pfQUYoP/aeHzGBgvBzb7uPyEi5vaWSnua0rpKcTmN/tDjs0ujFS32X2Ls6NS0QmMxWi2IqF+/vK73XH5/Y5CnMKioIw8IxUV394LNFuRUN/thuj+FIxFXj+tw+AjL/AW+AkeDv+RiotHlCWUyrfwtLtHe175iJKhEKfQRM2bvcCcB+bIxcW3uJNSlYqzENgAFtFIXpwthIlTijwHKEItNOBl5GVwyA/Z67Lv4rGi0IS4C2+hNKXi7EZwjOJW0djVYCf+nqTxIniCPoW5FivOTWggmD1SEGfzUfx3HNfSyEVxFgH7MG9/8XoDY2/i+JttrIh4l2m9u/UQZXAXmg7WiBBnzfTpojhbicZOyI2WQ+wTvHGKfi7TO7tdnLWB91WNliO8LWymF6cL/cwIIx3gA1WzLkw2Ptn4sd6bfy0mZoCRgjh7BE3+IjVdQPwD4nOwA7ShvnVnt+HLyFIwLDGfpoOVuPJw9hYaP180wVkORCjCEurfJPWzlKZmmMRjP2ffkpl7aPQKFWdXgSFQ8Z3/PyTG/gXUyxrtfPmmmgAAAABJRU5ErkJggg==" alt="Save & sync" /></button>
                        <button class="save">
                            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAFBhaW50Lk5FVCB2My41Ljg3O4BdAAADy0lEQVRoQ+3ZeUgUURwHcMvSwkrLzISy7BSKCIlEIiq3pMi0g+wCj1ELTIouw6DDysr6o6CCoH8iyiOtjA7tUmsroqg/isCgiMwgOzYtj1Lr9f2t++T1nF2PnYk15gsfWIe38953j9mZ0c2IESOdTm84AW/hvQuqhELwA4dRgHUDR8FhMoAGboMBLmgh0PrOg8PwIhutf7leIsAo4koxiqjFH1bDTogDH/hX0axIPNQCjeE+w1z4F9GkyAz4BWIJrg7Gg97RpMgVEBcvOwZ6R5MiFSAuXGYGvaNJkWcgLlx2GfSOJkXoKCUuXEYHAr2jSREveAzi4jl6N9xB72hShNIP9sEbqIdySAM6/e90ku81r0m+/7tp9QPGCB7/Ukrqtpj2nutlGyJHsyJinH4HsPhUXoJTSmq3mTKy7b0wuhRxOkYRo4hOUSuy5NTTI/9FkelpJ8z/RZGQhO1fQtce9LANkdN9ikxauZl59POeYhsix3WKJNz6Fp9Y1nAVimKLPlfIReJv1ED1d/w4Viil9VttT+NxqSJjku41m+UCsiRzk1m5XRtoexqP7kUGwqCWh+0noaR2wsrCyiq1AiQmt5yFph6Ksg0Xo2uRcUDXKh8gmDZ0IKMCQmb+WFVY2aZE3I1qFjgtktaQ3jL0r+hWJAw+Aj2HvIPR0F4u9XDvxYKjkllssYVOFq0llNIGFqLsZO6efWlfj1qG/hVdiiwCOgPmJTg6M5Y/22Lmg3Wsh5c3m5qSZS2QeLeJ4ayXefkP5/uh+wMjQIzmRVKhGfikspcQAHL6wCtoHdvXdyibteM0iz75kPmOnSzug6wHMZoV6QkHgU/0W3jcyn9iGH7YdlQty3u5P/bapxRs49kObcb38fFj/QOC2myHO9Aa0568o5NWbGKha7NeLC94bbJtVo2jIp6QDfJkbdDnnH95k+420jtACQK6baT6HDvoXacbgtbE36ypbD0oXLfQC2o39orQ3cQSkCdSJRZZUfCmAdvoXwIXxDGdkAzWOFuEvrjPQZ7ALrFITE45baNLYtWPYQcUgTXOFBkGVSDv3CGVIs74CYPBqSILQN5xuzQuQuj75VSRHjAL6CY1Nw+iJYthKTfvcHEhn1AprfsWceDiGTInsyBHNDszP9e0OzdfFJ6RfT5819kLorhiSzr2tUEpa6juapEuJbb4SxafUC8okmWbTjXrgIrQL7alq3BNUa82uZYSbn2n+2l20x/oMNvVI4yV15DhLPJ4GUu806i6CGfQPuOuf32CIiMxV7uha2ZXRd9bI0aMdDhubn8Ar6TpT8RF7mYAAAAASUVORK5CYII=" alt="Get as image" /></button>
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
    this._canvas.addEventListener('mousedown', this._getInitialCursorPosition, false);
    this._canvas.addEventListener('mouseup', this._draw, false);

    let pencilColors = document.querySelectorAll('.' + styles.pencil);
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
        pnp.sp.web.lists.getByTitle(this.properties.listTitle).items.add({
          'WebPartId': this.properties.drawId,
          'UserId': this._currentUser,
          'Layers': JSON.stringify(this._userDrawings)
        }).then((iar: ItemAddResult) => {
          console.log(iar);
        });
      } else {
        pnp.sp.web.lists.getByTitle(this.properties.listTitle).items.getById(this._spItemID).update({
          'Layers': JSON.stringify(this._userDrawings)
        }).then(i => {
          console.log(i);
        });
      }
    }
    //on récupère les dessins
    this._retrieveDrawings();
  }

  private _retrieveDrawings = () => {
    pnp.sp.web.lists.getByTitle(this.properties.listTitle).items.filter("WebPartID eq '" + this.properties.drawId + "'").get().then((items: any[]) => {
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
    });
  }

  //#endregion DataAccess
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
    this._pencilSize = event.target.value;
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
    let imageUrl = this._canvas.toDataURL("image/png");
    let image = document.querySelector("img#image");
    image.setAttribute('src', imageUrl);
    image.setAttribute('style', 'display:block;');
  }

  private _getInitialCursorPosition = (event) => {
    let pos = this._getMousePos(event);

    this._initialPosition = {
      'posX': pos.x,
      'posY': pos.y
    };
  }

  private _getMousePos = (event) => {
    let rect = this._canvas.getBoundingClientRect(), // abs. size of element
      scaleX = this._canvas.width / rect.width, // relationship bitmap vs. element for X
      scaleY = this._canvas.height / rect.height; // relationship bitmap vs. element for Y

    return {
      x: (event.clientX - rect.left) * scaleX, // scale mouse coordinates after they have
      y: (event.clientY - rect.top) * scaleY // been adjusted to be relative to element
    }
  }

  private _draw = (event) => {
    let layerName = this._generateGUID();
    let pos = this._getMousePos(event);
    this._finalPosition = {
      'posX': pos.x,
      'posY': pos.y
    };

    let drawing = {
      'layerName': layerName,
      'initialPosition': this._initialPosition,
      'finalPosition': this._finalPosition,
      'penShape': this._penShape,
      'pencilColor': this._pencilColor,
      'pencilSize': this._pencilSize,
      'penBoundary': this._penBoundary,
      'date': new Date(),
      'visible': true
    };

    /*
    'layerName' : ID de l'action
    'initialPosition' : Position initiale (au clic)
    'finalPosition' : Position finale (au relâchement)
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
    this._context2D.moveTo(drawing.initialPosition.posX, drawing.initialPosition.posY);
    this._context2D.lineTo(drawing.finalPosition.posX, drawing.finalPosition.posY);
    this._context2D.stroke();
  }

  private _drawRectangle = (drawing) => {
    let pos = {
      'posX': 0,
      'posY': 0
    };
    let width = 0;
    let height = 0;
    if (drawing.initialPosition.posX > drawing.finalPosition.posX) {
      pos.posX = drawing.finalPosition.posX;
      width = drawing.initialPosition.posX - drawing.finalPosition.posX;
    } else {
      pos.posX = drawing.initialPosition.posX;
      width = drawing.finalPosition.posX - drawing.initialPosition.posX;
    }
    if (drawing.initialPosition.posY > drawing.finalPosition.posY) {
      pos.posY = drawing.finalPosition.posY;
      height = drawing.initialPosition.posY - drawing.finalPosition.posY;
    } else {
      pos.posY = drawing.initialPosition.posY;
      height = drawing.finalPosition.posY - drawing.initialPosition.posY;
    }
    this._context2D.strokeRect(pos.posX, pos.posY, width, height);
  }

  private _drawCircle = (drawing) => {
    let pos = {
      'posX': 0,
      'posY': 0
    };
    var radius = 0;
    if (drawing.initialPosition.posX > drawing.finalPosition.posX) {
      radius = (drawing.initialPosition.posX - drawing.finalPosition.posX) / 2;
      pos.posX = drawing.finalPosition.posX + radius;
    } else {
      radius = drawing.finalPosition.posX - drawing.initialPosition.posX;
      pos.posX = drawing.initialPosition.posX + radius;
    }
    if (drawing.initialPosition.posY > drawing.finalPosition.posY) {
      pos.posY = drawing.finalPosition.posY + radius;
    } else {
      pos.posY = drawing.initialPosition.posY + radius;
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
