{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/03/2003
Modifié le ... :   /  /    
Description .. : Remplacé en eagl par CPTVAMODIF_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit TvaModif;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte, Mask, Ent1,
  HEnt1, Saisie, Saisutil, SaisBase, hmsgbox, HTB97, ColMemo, HPanel, UiUtil,
  HRichOLE, ADODB ;

Procedure TvaModifEnc ;

type
  TFTvaModif = class(TFMul)
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
    E_ECRANOUVEAU: THCritMaskEdit;
    XX_WHERE3: TEdit;
    E_ECHE: THCritMaskEdit;
    BZoomPiece: TToolbarButton97;
    HM: THMsgBox;
    XX_WHEREMODE: TEdit;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
  private
  UpSQL : String ;
  NowFutur : TDateTime ;
  Function  EstCeClient ( Gene,Auxi : String ) : boolean ;
  procedure RemplirOM ( OM : TMOD ; Q1 : TQuery ) ;
  procedure MajModTvaEnc ;
  public
  end;

implementation

{$R *.DFM}

Uses UtilPgi ;

Procedure TvaModifEnc ;
Var X : TFTvaModif ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
if _Blocage(['nrCloture','nrBatch','nrLettrage','nrEnca','nrDeca'],True,'nrSaisieModif') then Exit ;
X:=TFTvaModif.Create(Application) ;
X.FNomFiltre:='TVAMODIF' ; X.Q.Manuel:=TRUE ; X.Q.Liste:='TVAMODIF' ;
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



procedure TFTvaModif.FormShow(Sender: TObject);
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

procedure TFTvaModif.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFTvaModif.BZoomPieceClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisie(Q,taConsult,E_QUALIFPIECE.Text) ;
end;

Function TFTvaModif.EstCeClient ( Gene,Auxi : String ) : boolean ;
Var Q2 : TQuery ;
BEGIN
Result:=False ;
if Auxi<>'' then
   BEGIN
   Q2:=OpenSQL('Select T_AUXILIAIRE from TIERS Where T_AUXILIAIRE="'+Auxi+'" AND (T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD")',True) ;
   if Not Q2.EOF then Result:=True ;
   Ferme(Q2) ;
   END else
   BEGIN
   Q2:=OpenSQL('Select G_GENERAL from GENERAUX Where G_GENERAL="'+Gene+'" AND G_NATUREGENE="TID"',True) ;
   if Not Q2.EOF then Result:=True ;
   Ferme(Q2) ;
   END ;
END ;

procedure TFTvaModif.RemplirOM ( OM : TMOD ; Q1 : TQuery ) ;
Var k : integer ;
BEGIN
FillChar(OM.MODR,Sizeof(OM.MODR),#0) ;
OM.ModR.Action:=taModif ;
OM.ModR.NbEche:=1 ;
OM.ModR.TotalAPayerP:=Q1.FindField('E_DEBIT').AsFloat+Q1.FindField('E_CREDIT').AsFloat ;
OM.ModR.TotalAPayerD:=Q1.FindField('E_DEBITDEV').AsFloat+Q1.FindField('E_CREDITDEV').AsFloat ;
OM.ModR.CodeDevise:=V_PGI.DevisePivot ;
OM.ModR.TauxDevise:=1.0 ;
OM.ModR.Quotite:=1 ;
OM.ModR.Decimale:=V_PGI.OkDecV ;
OM.ModR.DateFact:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ;
OM.ModR.DateBL:=OM.MODR.DateFact ;
// GP REGL
OM.ModR.DateFactExt:=Q1.FindField('E_DATEREFEXTERNE').AsDateTime ;
If OM.ModR.DateFactExt=IDate1900 Then OM.ModR.DateFactExt:=OM.ModR.DateFact ;
OM.ModR.Aux:=Q1.FindField('E_AUXILIAIRE').AsString ;
OM.ModR.ModifTva:=False ;
OM.ModR.TabEche[1].ModePaie:=Q1.FindField('E_MODEPAIE').AsString ;
OM.ModR.TabEche[1].DateEche:=Q1.FindField('E_DATEECHEANCE').AsDateTime ;
OM.ModR.TabEche[1].MontantP:=OM.MODR.TotalAPayerP ;
OM.ModR.TabEche[1].MontantD:=OM.MODR.TotalAPayerD ;
OM.ModR.TabEche[1].Pourc:=100.0 ;
{#TAVENC}
for k:=1 to 4 do OM.MODR.TabEche[1].TAV[k]:=Q1.FindField('E_ECHEENC'+IntToStr(k)).AsFloat ;
OM.ModR.TabEche[1].TAV[5]:=Q1.FindField('E_ECHEDEBIT').AsFloat ;
END ;

procedure TFTvaModif.MajModTvaEnc ;
BEGIN
if ExecuteSQL(UpSQL)<>1 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFTvaModif.FListeDblClick(Sender: TObject);
Var Q1 : TQuery ;
    SQL,Regime,Nature,CodeTva,StW : String ;
    OM  : TMOD ;
    LastDate : TDateTime ;
    Client : boolean ;
    RTVA   : Enr_Base ;
begin
  inherited;
if ((Q.EOF) And (Q.Bof)) then Exit ;
OM:=Nil ; UpSQL:='' ; LastDate:=0 ;
StW:='where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
    +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
    +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
    +' AND E_NUMLIGNE='+IntToStr(Q.FindField('E_NUMLIGNE').AsInteger)
    +' AND E_NUMECHE='+IntToStr(Q.FindField('E_NUMECHE').AsInteger)
    +' AND E_ECHE="X" and E_NUMECHE>0 AND E_QUALIFPIECE="N"'
    +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString ;
SQL:='Select * from Ecriture '+StW ;
Q1:=OpenSQL(SQL,True) ; Client:=False ;
if Not Q1.EOF then
   BEGIN
   Regime:=Q1.FindField('E_REGIMETVA').AsString ;
   Nature:=Q1.FindField('E_NATUREPIECE').AsString ;
   CodeTva:=Q1.FindField('E_TVA').AsString ;
   LastDate:=Q1.FindField('E_DATEMODIF').AsDateTime ;
   Client:=EstCeClient(Q1.FindField('E_GENERAL').AsString,Q1.FindField('E_AUXILIAIRE').AsString) ;
   OM:=TMOD.Create ;
   RemplirOM(OM,Q1) ;
   END ;
Ferme(Q1) ;
if OM<>Nil then
   BEGIN
   RTVA.Regime:=Regime ; RTVA.Client:=Client ; RTVA.Action:=taModif ;
   RTVA.Nature:=Nature ; RTVA.CodeTva:=CodeTva ;
   if SaisieBasesHT(OM,RTVA) then
      BEGIN
      if OM.MODR.ModifTva then
         BEGIN
         UpSQL:='UPDATE ECRITURE SET E_ECHEENC1='+StrfPoint(OM.MODR.TabEche[1].TAV[1])
               +', E_ECHEENC2='+StrfPoint(OM.MODR.TabEche[1].TAV[2])
               +', E_ECHEENC3='+StrfPoint(OM.MODR.TabEche[1].TAV[3])
               +', E_ECHEENC4='+StrfPoint(OM.MODR.TabEche[1].TAV[4])
               +', E_ECHEDEBIT='+StrfPoint(OM.MODR.TabEche[1].TAV[5])
               +', E_TVA="'+RTVA.CodeTva+'"'
               +', E_DATEMODIF="'+USTime(NowFutur)+'"'
               +StW+' AND E_DATEMODIF="'+USTime(LastDate)+'"' ;
         if Transactions(MajModTvaEnc,3)<>oeOk then MessageAlerte(HM.Mess[0]) ;
         END ;
      END ;
   END ;
OM.Free ;
end;

end.
