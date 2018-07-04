unit TvaValid ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB, Hqry, StdCtrls, Grids, DBGrids, HDB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte, Mask, Ent1,
  HEnt1, Saisie, Saisutil, SaisBase, hmsgbox, HStatus, HTB97, ColMemo, HPanel, UiUtil,
  HRichOLE, ADODB ;

Procedure TvaValidEnc ;

type
  TFTvaValid = class(TFMul)
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    TE_GENERAL: THLabel;
    E_GENERAL: THCpteEdit;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    TE_DATECREATION: THLabel;
    E_DATECREATION: THCritMaskEdit;
    TE_DATECREATION_: THLabel;
    E_DATECREATION_: THCritMaskEdit;
    TE_DATEECHEANCE: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    TE_DATEECHEANCE2: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_REFINTERNE: THLabel;
    E_REFINTERNE: TEdit;
    TE_MODEPAIE: THLabel;
    E_MODEPAIE: THValComboBox;
    TE_REGIMETVA: THLabel;
    E_REGIMETVA: THValComboBox;
    TE_LIBELLE: THLabel;
    E_LIBELLE: TEdit;
    E_NUMECHE: THCritMaskEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHERE3: TEdit;
    E_ECHE: THCritMaskEdit;
    BZoomPiece: TToolbarButton97;
    HM: THMsgBox;
    E_EDITEETATTVA: THCritMaskEdit;
    XX_WHEREAN: TEdit;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure ValEditEnc ;
    procedure UpdateValidEnc ;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
  NowFutur : TDateTime ;
  public
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  UtilPgi ;

Procedure TvaValidEnc ;
Var X  : TFTvaValid ;
    PP : THPanel ;
BEGIN
// GP le 08/06/99 if PasCreerDate(V_PGI.DateEntree) then Exit ;
if Not VH^.OuiTvaEnc then BEGIN HShowMessage('10;Tva sur encaissements;Le module n''est pas installé !;E;O;O;O;','','') ; Exit ; END ;
if _Blocage(['nrCloture','nrBatch','nrLettrage','nrEnca','nrDeca'],True,'nrSaisieModif') then Exit ;
X:=TFTvaValid.Create(Application) ;
X.FNomFiltre:='TVAVALID' ; X.Q.Manuel:=TRUE ; X.Q.Liste:='TVAVALID' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
     _Bloqueur('nrSaisieModif',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFTvaValid.FormShow(Sender: TObject);
begin
E_DEVISE.Value:=V_PGI.DevisePivot ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
   E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
E_DATECREATION.Text:=StDate1900 ; E_DATECREATION_.Text:=StDate2099 ;
PositionneEtabUser(E_ETABLISSEMENT) ;
if ((E_JOURNAL.Value='') and (E_JOURNAL.Values.Count>0)) then
   BEGIN
   if Not E_JOURNAL.Vide then E_JOURNAL.Value:=E_JOURNAL.Values[0] else
    if E_JOURNAL.Values.Count>1 then E_JOURNAL.Value:=E_JOURNAL.Values[1] ;
   END ;
  inherited;
end;

procedure TFTvaValid.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFTvaValid.BZoomPieceClick(Sender: TObject);
begin
  inherited;
if ((Q.EOF) And (Q.BOF)) then Exit ;
TrouveEtLanceSaisie(Q,taConsult,E_QUALIFPIECE.Text) ;
end;

procedure TFTvaValid.UpdateValidEnc ;
Var StW : String ;
    Nb  : Integer ;
BEGIN
StW:='E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
    +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
    +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
    +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
    +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
    +' AND E_NUMECHE>=1 AND E_QUALIFPIECE="N" AND E_EDITEETATTVA="X"' ;
Nb:=ExecuteSQL('UPDATE ECRITURE SET E_EDITEETATTVA="#", E_DATEMODIF="'+USTime(NowFutur)+'" WHERE '+StW) ;
if Nb<=0 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFTvaValid.ValEditEnc ;
Var i : integer ;
BEGIN
if Fliste.AllSelected then
   BEGIN
   InitMove(100,'') ;
   Q.First ;
   While Not Q.EOF do
      BEGIN
      MoveCur(False) ;
      UpdateValidEnc ; if V_PGI.IoError<>oeOk then Break ;
      Q.Next ;
      END ;
   END else
   BEGIN
   InitMove(FListe.NbSelected,'') ;
   for i:=0 to FListe.NbSelected-1 do
       BEGIN
       MoveCur(False) ;
       FListe.GotoLeBookmark(i) ;
       UpdateValidEnc ; if V_PGI.IoError<>oeOk then Break ;
       END ;
   END ;
FiniMove ;
END ;

procedure TFTvaValid.BOuvrirClick(Sender: TObject);
begin
  inherited;
if ((Not FListe.AllSelected) and (FListe.NbSelected<=0)) then BEGIN HM.Execute(2,caption,'') ; Exit ; END ;
if HM.Execute(1,caption,'')<>mrYes then Exit ;
NowFutur:=NowH ;
if Transactions(ValEditEnc,2)<>oeOk then MessageAlerte(HM.Mess[0]) else
   BEGIN
   if Not FListe.AllSelected then Fliste.ClearSelected else FListe.AllSelected:=False ;
   BChercheClick(Nil) ; 
   END ;
end ;

procedure TFTvaValid.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then _Bloqueur('nrSaisieModif',False) ;
  inherited;
end;

end.
