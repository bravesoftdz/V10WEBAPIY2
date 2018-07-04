{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : 22/04/2003 : mise à jour correcte de SO_EXOV8 ( FQ
Suite ........ : 10198 )
Suite ........ : 22/04/2003 - CA - FQ 12294 - Mise à jour de SO_CPEXOREF
Suite ........ : 22/04/2003 - CA - FQ 10193 - annulation clôture provisoire
Suite ........ : 20/01/2004 - CA - FQ 13202 - Mise à jour des infos de clôture immo
Mots clefs ... : CLOTURE;SO_EXOV8
*****************************************************************}
unit AnulClo;

interface

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  hmsgbox, StdCtrls, Hctrls, Buttons, ExtCtrls,
  ENT1, HEnt1, HSysMenu, Mask, SoldeCpt, HPanel, UiUtil, HTB97 , ParamSoc,
  ADODB;

Function AnnuleclotureComptable(Definitive : Boolean ; FromParam : boolean = false ) : Boolean ;

type
  TFDECLO = class(TForm)
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    GBFermeDEF: TGroupBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    Confirmation: THMsgBox;
    Q: TQuery;
    HLabel6: THLabel;
    HMTrad: THSystemMenu;
    GBFERMEPRO: TGroupBox;
    HLabel7: THLabel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    HLabel11: THLabel;
    HLabel12: THLabel;
    GroupBox1: TGroupBox;
    HLabel10: THLabel;
    DateDebN1: TMaskEdit;
    Label7: TLabel;
    DateFinN1: TMaskEdit;
    HLabel13: THLabel;
    DateDebN: TMaskEdit;
    Label9: TLabel;
    DateFinN: TMaskEdit;
    HLabel14: THLabel;
    DateDebM: TMaskEdit;
    Label1: TLabel;
    DateFinM: TMaskEdit;
    Label2: TLabel;
    Label3: TLabel;
    Dock: TDock97;
    HPatienter: THLabel;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    Definitive : Boolean ;
    Automatique : Boolean ;
    Exo1 : TExoDate ;
    Exo2 : TExoDate ;
    OnSort : Boolean ;
    function DeleteANO : Integer ;
    function DeleteCLO : Integer ;
    procedure MajExo ( bDefinitive : boolean )  ;
    function DeleteANOH: Integer;
    procedure MajExoV8;
    function PresenceLettragePointage: boolean;
  public
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}
  CpteUtil,
  CpteSav,
  LicUtil,
  UtilPgi,
  uLibCloture,
  uLibRevision; // OuvertureExerciceRIC;

Function AnnuleclotureComptable(Definitive : Boolean ; FromParam : boolean = false) : Boolean ;
var FDeClo: TFDEClo;
    OutProg : Boolean ;
    PP : THPanel ;
begin
Result:=FALSE ; if Not _BlocageMonoPoste(True) then Exit ;
if not Definitive and GetParamSocSecur('SO_CPANODYNA', false) and not FromParam then
 begin
  PGIInfo('La gestion des A-nouveaux dynamiques est activée, désactivez celle ci avant d''annuler la cloture provisoire') ;
  exit ;
 end ;
FDEClo:=TFDEClo.Create(Application) ; OutProg:=FALSE ;
FDEClo.Definitive:=Definitive ; FDeclo.Automatique:=FALSE ; FDeclo.OnSort:=OutProg ;
if Not Definitive then FDeclo.HelpContext:=7742200 ;
PP:=FindInsidePanel ;
if (PP=Nil) then
   BEGIN
    try
     FDEClo.ShowModal ;
    Finally
     OutProg:=FDeclo.OnSort ;
     FDEClo.free ;
     _DeblocageMonoPoste(True) ;
    End ;
   END else
   BEGIN
   InitInside(FDEClo,PP) ;
   FDEClo.Show ;
   END ;
If OutProg Then
 Result:=TRUE ;
Screen.Cursor:=SyncrDefault ;
end ;


function TFDECLO.DeleteANO : Integer ;
begin
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('DELETE FROM ECRITURE WHERE E_EXERCICE="'+EXO2.Code+'" And E_ECRANOUVEAU="OAN"') ;
ChangeSQL(Q) ; Q.ExecSQL ;
Result:=Q.rowsaffected ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+EXO2.Code+'" And Y_ECRANOUVEAU="OAN"') ;
ChangeSQL(Q) ; Q.ExecSQL ;
end ;

function TFDECLO.DeleteCLO : Integer ;
begin
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('DELETE FROM ECRITURE WHERE E_EXERCICE="'+EXO1.Code+'" And E_ECRANOUVEAU="C"') ;
ChangeSQL(Q) ; Q.ExecSQL ;
Result:=Q.rowsaffected ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+EXO1.Code+'" And Y_ECRANOUVEAU="C"') ;
ChangeSQL(Q) ; Q.ExecSQL ;
end ;

procedure TFDECLO.MajExo ( bDefinitive : boolean ) ;
BEGIN
{Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_EXERCICE="'+VH^.Precedent.Code+'"') ;
ChangeSQL(Q) ; Q.ExecSQL ;
If VH^.Suivant.Code<>'' Then
   BEGIN
   Q.Close ; Q.SQL.Clear ;
   Q.SQL.Add('UPDATE EXERCICE SET EX_ETATCPTA="NON" WHERE EX_EXERCICE="'+VH^.Suivant.Code+'"') ;
   ChangeSQL(Q) ; Q.ExecSQL ;
   END ;}
   if bDefinitive then
   begin
    ExecuteSQL ('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_EXERCICE="'+VH^.Precedent.Code+'"');
    If VH^.Suivant.Code<>'' Then
      ExecuteSQL ('UPDATE EXERCICE SET EX_ETATCPTA="NON" WHERE EX_EXERCICE="'+VH^.Suivant.Code+'"') ;
   end else ExecuteSQL ('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_EXERCICE="'+VH^.EnCours.Code+'"');
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2006
Modifié le ... :   /  /
Description .. : - LG - FB 10678 - qd on annule la cloture provisoire on
Suite ........ :  remets le paramsoc en l'etat
Mots clefs ... :
*****************************************************************}
procedure TFDECLO.BValiderClick(Sender: TObject);
var i,NbDeleteANO,vAnnulClo : Integer ;
    OnArrete : Boolean ;
    St : String ;
    OkDec : Boolean;
begin
If Definitive Then
  BEGIN
{$IFNDEF CERTIFNF}
  If GetParamSocSecur('SO_CPCONFORMEBOI', False) then  //06/12/2006 YMO Norme NF 203
  begin
   PgiInfo('Pour la conformité stricte avec la norme NF 203 (et le BOI du 24/01/2006) cette fonction n''est plus disponible',Caption);
   Exit;
  end;
{$ENDIF}

//  If (not (ctxPCL in V_PGI.PGIContexte)) and (V_PGI.PassWord<>CryptageSt(DayPass(Date))) Then
  If (not (ctxPCL in V_PGI.PGIContexte)) and ( not V_PGI.FSuperviseur ) Then
    BEGIN
    St:='10;'+Confirmation.Mess[8]+';'+Confirmation.Mess[9]+#13+#10+Confirmation.Mess[10] ;
    Confirmation.Mess[11]:=St ;
    Confirmation.Execute(11,'','') ;
    Exit ;
    END ;
  END ;
OnArrete:=FALSE ;
{$IFNDEF EURO}
If VH^.Entree.Code<>VH^.EnCours.Code Then
   BEGIN
   Confirmation.Execute(3,'','') ; Exit ;
   END ;
{$ENDIF}
{ Ajout CA le 18/04/2001 }
//if ctxPCL in V_PGI.PGIContexte then
if Definitive then  // CA - 22/04/2003 - FQ 10193
begin
  if ((VH^.Precedent.Code <> '') and (not ExisteSQL('SELECT * FROM ECRITURE WHERE E_EXERCICE="'+VH^.Precedent.Code+'"'))) then
  begin
    PGIInfo('Suppression du journal d''à-nouveaux impossible. #10#13L''exercice précédent n''a pas d''écriture.',Caption);
    Exit;
  end;

  PGIInfo('En référence au BOI 13 L-1-06 N° 12 du 24 janvier 2006 paragraphe 29, les écritures ' + #13#10 +
          'non validées concernant cet exercice vont être validées lors de la dé clôture.' + #13#10 +
          'Seul l''ajout d''écritures complémentaires est autorisé.', Caption);

end;
{ Fin Ajout CA le 18/04/2001 }

// ajout me 12/09/2002
OkDec := FALSE;
if (GetParamSocSecur('SO_CPLIENGAMME','') = 'S1')
and (ExisteSQL('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" and E_EXERCICE="'+VH^.Encours.Code+'"')) then
begin
    if (PGIAsk('Attention, des à nouveaux lettrable et/ou pointable sont présents, ainsi qu''un lien avec un produit PGI.'+#10#13+' L''export de type synchronisation sera impossible.'+#10#13
    +' Voulez-vous continuez ?','Annulation de clôture'))<> mrYes then
    exit
    else
    OkDec := TRUE;
end;

If Definitive and (VH^.Suivant.EtatCpta='OUV') then // 14022
  Confirmation.Execute(12,'',''); // Le deuxième exercice ouvert nécessitera une réouverture manuelle.

If Definitive then BEGIN Exo2:=VH^.EnCours ; Exo1:=VH^.Precedent ; END
              Else BEGIN Exo2:=VH^.Suivant ; Exo1:=VH^.EnCours ; END ;

              i:=Confirmation.Execute(0,'','') ;
If i<>mrYes Then Exit ; Screen.Cursor:=SynCrDefault ;
i:=Confirmation.Execute(1,'','') ;
If i<>mrYes Then Exit ; Screen.Cursor:=SynCrDefault ;
HPatienter.Visible:=TRUE ;
EnableControls(Self,False) ;
{$IFDEF EURO}
EcritProcEuro('Annulation clôture provisoire') ;
{$ENDIF}
Try
  BeginTrans ;
  NbDeleteANO:=DeleteANO ;
  { Ajout CA le 09/03/2001 }
  if ((NbDeleteANO=0) and (ctxPCL in V_PGI.PGIContexte)) then
  begin
    if not PresenceLettragePointage then
      NbDeleteANO := DeleteANOH  // Suppression des à-nouveaux typés 'H'
    else
    begin
      PGIInfo('Clôture impossible. #10#13Vous devez supprimer auparavant le lettrage et le pointage sur les à-nouveaux.',Caption);
      OnArrete := True;
    end;
  end;
  { Fin CA le 09/03/2001 }
  { Modif CA le 09/03/2001 }
  //If NbDeleteANO<=0 Then BEGIN Confirmation.Execute(2,'','') ; OnArrete:=TRUE ;  END ;
  if (not(ctxPCL in V_PGI.PGIContexte) and (NbDeleteANO<=0)) Then
  begin
    Confirmation.Execute(2,'','') ;
    OnArrete:=TRUE ;
  end;
  // Fin CA le 09/03/2001

  if not OnArrete Then
  begin
    DeleteClo ;
    if Definitive Then
    begin
      // Mise à jour de SO_EXOV8 CA - 22/04/2003
      MajExoV8;
      MajExo (Definitive);
      ChargeMagExo(False) ;
      // CA - 22/04/2003 - Mise à jour de SO_CPEXOREF
      SetParamSoc('SO_CPEXOREF', VH^.EnCours.Code);
      // GCO - 20/07/2005 - FQ 15025
      SetParamSoc('SO_CPDERNEXOCLO', VH^.Precedent.Code);

      // GCO - 19/09/2006 - Validations automatiques des écritures concernées par la décloture
      { BVE 27.08.07
      On ne valide plus les ecritures sur la décloture }
{$IFNDEF CERTIFNF}
      ExecuteSQL('UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + VH^.EnCours.Code + '" ' +
                 'AND E_DATECOMPTABLE>="' + USDateTime(VH^.EnCours.Deb) + '" ' +
                 'AND E_DATECOMPTABLE<="' + USDateTime(VH^.EnCours.Fin) + '" ' +
                 'AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N"');
{$ENDIF}

      // GCO - 04/09/2006 - FQ 18662
      MajValidationJournaux( False );

      RecalculSouchePourCloture ;
      { Mise à jour des infos de clôture immo - CA - 20/01/2004 - FQ 13202 }
      if ((ExisteSQL('SELECT I_IMMO FROM IMMO')) and (GetParamSocSecur ('SO_EXOCLOIMMO','')='')) then
      begin
        SetParamSoc('SO_DATECLOTUREIMMO',Date);
        SetParamSoc('SO_EXOCLOIMMO',VH^.Encours.Code);
      end;
      OnSort:=TRUE ;
    end
    else // Provisoire
    begin
      MajExo (Definitive);
      OnSort:=TRUE ;
    end ;
    MajTotTousComptes(TRUE,'') ;

{$IFDEF EAGLCLIENT}
    AvertirCacheServer( 'PARAMSOC' ) ; // SBO : 30/06/2004
    AvertirCacheServer( 'EXERCICE' ) ; // SBO : 30/06/2004
{$ENDIF}
    CHARGESOCIETEHALLEY;
    AvertirMultiTable('TTEXERCICE');
  end;

  CommitTrans ;     

  if (Definitive) and (not OnArrete) then //YMO 11/12/2006 Seulement si le traitement est éxécuté
  begin
    // GCO - 18/09/2006 - Traçage des évènements       
    { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}    
    CPEnregistreLog('ANNULCLOEXO ' + VH^.EnCours.Code);
{$ELSE}
    CPEnregistreJalEvent('CEX','Clôture exercice','ANNULCLOEXO ' + VH^.EnCours.Code);
{$ENDIF}

    // CA - FQ 19855 - 02/05/2007 - Si premier exercice, on ne décrémente pas le compteur
    if (not Label2.Visible) then
    begin
      vAnnulClo:=GetParamSocSecur('SO_CPANNULCLO', False);
      //06/12/2006 YMO Norme NF 203 : interdiction d'annuler en cascade
      SetParamSoc('SO_CPANNULCLO',vAnnulClo+1);
    end;

    // GCO - 12/10/2007 - FQ 21633
    if VH^.Revision.Plan <> '' then
      SynchroRICAvecExercice;

  end;
Except
  OnSort:=FALSE ;
  ShowMessage(Confirmation.Mess[5]) ;
  Rollback ;
end ;

EnableControls(Self,TRUE) ;
HPatienter.Visible:=FALSE ;
BValider.Visible:=FALSE ;

// ajout me 12/09/2002
if OkDec then SetParamsoc ('SO_CPDECLOTUREDETAIL',TRUE);
if not ( Parent is THPanel ) then
 Close ;

end;

procedure TFDECLO.FormShow(Sender: TObject);
begin
GBFermeDef.Visible:=Definitive ;
GBFermePro.Visible:=Not Definitive ;
If Definitive Then Caption:=Confirmation.Mess[7] Else Caption:=Confirmation.Mess[6] ;
UpdateCaption(Self) ;
ExoToDates(VH^.EnCours.Code,DateDebN,DateFinN) ;
If VH^.Precedent.Code<>'' Then ExoToDates(VH^.Precedent.Code,DateDebM,DateFinM) Else
   BEGIN
   DateDebM.Visible:=FALSE ; DateFinM.Visible:=FALSE ; Label2.Visible:=TRUE ;
   END ;
If VH^.Suivant.Code<>'' Then ExoToDates(VH^.Suivant.Code,DateDebN1,DateFinN1)  Else
   BEGIN
   DateDebN1.Visible:=FALSE ; DateFinN1.Visible:=FALSE ; Label3.Visible:=TRUE ;
   END
end;

procedure TFDECLO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(True) ;
   Action:=caFree ;
   END ;
  if isInside(self) then CloseInsidePanel(Self);
end;

procedure TFDECLO.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{ Ajout CA le 09/03/2001 }
// Suppression des à-nouveaux typés "H"
function TFDECLO.DeleteANOH : Integer ;
{var QV8 : TQuery;
    Exo : string;}
begin
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('DELETE FROM ECRITURE WHERE E_EXERCICE="'+EXO2.Code+'" And E_ECRANOUVEAU="H"') ;
ChangeSQL(Q) ; Q.ExecSQL ;
Result:=Q.rowsaffected ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('DELETE FROM ANALYTIQ WHERE Y_EXERCICE="'+EXO2.Code+'" And Y_ECRANOUVEAU="H"') ;
ChangeSQL(Q) ; Q.ExecSQL ;
// Mise à jour de SO_EXOV8
{ QV8 := OpenSQL ('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" ORDER BY E_DATECOMPTABLE DESC',True);
  if not QV8.Eof then
  begin
    Exo := QV8.FindField('E_EXERCICE').AsString;
    SetParamSoc ('SO_EXOV8',Exo);
  end else SetParamSoc ('SO_EXOV8','');
  Ferme (QV8);}
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/04/2003
Modifié le ... : 22/04/2003
Description .. : Mise à jour du Paramsoc SO_EXOV8
Mots clefs ... : SO_EXOV8;DECLOTURE
*****************************************************************}
procedure TFDECLO.MajExoV8;
var QExo : TQuery;
    Exo , ExoV8 : string;
    bTrouve : boolean;
begin
  ExoV8 := '';  bTrouve := False;
  QExo := OpenSQL ( 'SELECT * FROM EXERCICE WHERE EX_DATEDEBUT<"'+USDateTime(VH^.EnCours.Deb)+'" ORDER BY EX_DATEDEBUT DESC',True );
  try
    while not QExo.Eof do
    begin
      Exo := QExo.FindField('EX_EXERCICE').AsString;
      { L'exercice a t'il des à nouveaux ? }
      if not ExisteSQL('SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND (E_ECRANOUVEAU="H" OR E_ECRANOUVEAU="OAN")') then
      begin
        ExoV8 := Exo;
        bTrouve := True;
      end
      { L'exercice a t'il des à nouveaux H? }
      else if ExisteSQL ( 'SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="H"') then
      begin
        ExoV8 := Exo;
        bTrouve := True;
      end
      { Une écriture est-elle lettrée sur les à-nouveaux OAN : suite récup. sisco II }
      else if ExisteSQL  ( 'SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN" AND E_LETTRAGE <>""') then
      begin
        { L' exercice à des à-nouveau OAN lettrés ==> on passe ces à-nouveaux à H }
        ExecuteSQL ( 'UPDATE ECRITURE SET E_ECRANOUVEAU="H" WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN"');
        ExecuteSQL ( 'UPDATE ANALYTIQ SET Y_ECRANOUVEAU="H" WHERE Y_EXERCICE="'+Exo+'" AND Y_ECRANOUVEAU="OAN"');
        ExoV8 := Exo;
        bTrouve := True;
      end;
      if bTrouve then break;
      QExo.Next;
    end;
  finally
    Ferme ( QExo );
  end;
  { si exoV8 n'est pas renseigné, dans le cas où le premier exercice est à OAN, on force les à-nouveaux à H }
  if ExoV8 = '' then
  begin
    QExo := OpenSQL ( 'SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',True);
    if not QExo.Eof then
    begin
      Exo := QExo.FindField('EX_EXERCICE').AsString;
      if ExisteSQL ('SELECT * FROM ECRITURE WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN"') then
      begin
        ExecuteSQL ('UPDATE ECRITURE SET E_ECRANOUVEAU="H" WHERE E_EXERCICE="'+Exo+'" AND E_ECRANOUVEAU="OAN"');
        ExecuteSQL ('UPDATE ANALYTIQ SET Y_ECRANOUVEAU="H" WHERE Y_EXERCICE="'+Exo+'" AND Y_ECRANOUVEAU="OAN"');
        ExoV8 := Exo;
      end;
    end;
    Ferme ( QExo );
  end;
  SetParamSoc ( 'SO_EXOV8', ExoV8 );
end;

function TFDECLO.PresenceLettragePointage : boolean;
begin
  Result := ExisteSQL('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" AND E_EXERCICE="'+
    EXO2.Code+'" AND ((E_ETATLETTRAGE<>"RI" AND E_ETATLETTRAGE<>"AL") OR (E_REFPOINTAGE<>""))');
end ;
{ Fin Ajout CA le 09/03/2001 }
procedure TFDECLO.BFermeClick(Sender: TObject);
begin
  //SG6 06/12/2004 FQ 14987

  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

end.
