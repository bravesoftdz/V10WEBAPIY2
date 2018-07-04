unit ImRapInt;

// CA - 17/05/1999 - Ajout du bouton Zoom
// mbo - 21/12/2006 - fq 19301 - modifier le bouton 'croix rouge' en 'mouette verte'

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  Grids,
  Hctrls,
  StdCtrls,
  Buttons,
  ExtCtrls,
  HSysMenu,
  hmsgbox,
  {$IFDEF EAGLCLIENT}
  utileAGL,
  uTob,
  {$ELSE}
  PrintDBG,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  HEnt1;

type
  TCRPiece = class
    NumPiece : longint;
    Journal : string;
    Date : TDateTime;
    Debit,Credit : double;
    NumLigne : integer;
    QualifPiece : string ;
  end;
  TFRappInteg = class(TForm)
    Pbouton: TPanel;
    TNumPiece: TLabel;
    Panel1: TPanel;
    BAide: TBitBtn;
    BFerme: TBitBtn;
    BImprimer: TBitBtn;
    Panel2: TPanel;
    BRechercher: TBitBtn;
    FListe: THGrid;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    FindDialog: TFindDialog;
    BZoomEcr: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    FindFirst : Boolean ;
    fCR : TList;
    FQualifPiece : string ;
  public
    { Déclarations publiques }
  end;

procedure AfficheCRIntegration (CR : TList);

implementation

{$IFDEF SERIE1}
{$ELSE}
uses Saisie;
{$ENDIF}

{$R *.DFM}

procedure AfficheCRIntegration (CR : TList);
var  FRappInteg: TFRappInteg;
begin
  FRappInteg := TFRappInteg.Create (Application);
  FRappInteg.fCR := CR;
  try
     FRappInteg.ShowModal;
  finally
     FRappInteg.FListe.VidePile (False);
     FRappInteg.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 05/11/2003
Modifié le ... :   /  /    
Description .. : LG - 05/11/2003 - j'ai pas toujours les montants de mvt. 
Suite ........ : donc -1 ds la debit on affiche aps l'info
Mots clefs ... : 
*****************************************************************}
procedure TFRappInteg.FormShow(Sender: TObject);
var i : integer;
    ALigneCR : TCRPiece;
begin
  FListe.ColAligns[0] := taRightJustify;
  FListe.ColAligns[1] := taCenter;
  FListe.ColAligns[2] := taCenter;
  FListe.ColAligns[3] := taRightJustify;
  FListe.ColAligns[4] := taRightJustify;
  FQualifPiece        := '' ;     
  if fCR.Count = 1 then TNumPiece.Caption := '1 ' + HM.Mess[0]
  else TNumPiece.Caption := IntToStr (fCR.Count) + ' ' + HM.Mess[1];
  FListe.RowCount := fCR.Count + 1;
  for i:=0 to fCR.Count - 1 do
  begin
    ALigneCR := fCR.Items[i];
    FListe.Cells[0,i+1] := IntToStr(ALigneCR.NumPiece);
    FListe.Cells[1,i+1] := ALigneCR.Journal;
    FListe.Cells[2,i+1] := DateToStr(ALigneCR.Date);
    if ALigneCR.Debit<>-1 then
    FListe.Cells[3,i+1] := StrfMontant(ALigneCR.Debit,15,V_PGI.OkDecV,'',True);
    if ALigneCR.Credit<>-1 then
    FListe.Cells[4,i+1] := StrfMontant(ALigneCR.Credit,15,V_PGI.OkDecV,'',True);
    FQualifPiece := ALigneCR.QualifPiece ;
  end;
  TNumPiece.Visible := True;
end;

procedure TFRappInteg.BImprimerClick(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
  PrintDBGrid (Caption,FListe.Name,'');
  {$ELSE}
  PrintDBGrid (FListe,Nil, Caption,'');
  {$ENDIF}
end;

procedure TFRappInteg.BRechercherClick(Sender: TObject);
begin FindFirst:=True; FindDialog.Execute ;end;

procedure TFRappInteg.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FindFirst); end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 22/10/2003
Modifié le ... :   /  /    
Description .. : - LG - 22/10/2003 - FB 12964 - le select etait fait 
Suite ........ : uniquement sur journal et numero
Mots clefs ... : 
*****************************************************************}
procedure TFRappInteg.FListeDblClick(Sender: TObject);
var Q : TQuery;
Year, Month, Day : Word ;
Cols, Where, OrderBy : string ;
begin
 DecodeDate(StrToDate(FListe.Cells[2,FListe.Row]), Year, Month, Day) ;
 Cols:='E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE' ;
 Where:='E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'"'
        +' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(Year, Month, 1)))+'"'
        +' AND E_JOURNAL="'+FListe.Cells[1,FListe.Row]+'"'
        +' AND E_QUALIFPIECE="'+FQualifPiece+'"'
        +' AND E_NUMEROPIECE='+FListe.Cells[0,FListe.Row]
        +' AND E_NUMLIGNE=1 AND E_NUMECHE<=1' ;
 OrderBy:='E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, '
          +'E_NUMECHE, E_QUALIFPIECE' ;
 Q:=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE '+Where+' ORDER BY '+OrderBy, TRUE) ;
{$IFDEF SERIE1}
  // A FAIRE
  {$ELSE}
  if not TrouveEtLanceSaisie(Q, taConsult,FQualifPiece) then HM.Execute (2,Caption,'');
  {$ENDIF}
  Ferme (Q);
end;

procedure TFRappInteg.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
