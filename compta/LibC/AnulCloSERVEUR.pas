{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/04/2003
Modifié le ... : 17/11/2003
Description .. : 22/04/2003 : mise à jour correcte de SO_EXOV8 ( FQ
Suite ........ : 10198 )
Suite ........ : 22/04/2003 - CA - FQ 12294 - Mise à jour de 
Suite ........ : SO_CPEXOREF
Suite ........ : 22/04/2003 - CA - FQ 10193 - annulation clôture provisoire
Suite ........ : 
Suite ........ : 11/07/2003 - SBO - Modification pour utilisation eAGL avec 
Suite ........ : process Serveur
Mots clefs ... : CLOTURE;SO_EXOV8
*****************************************************************}
unit AnulCloSERVEUR;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, StdCtrls, Hctrls, Buttons, ExtCtrls,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  DB, DBTables,
  {$ENDIF}
  uTOB,
  ENT1,
  HEnt1,
  HSysMenu,
  Mask,
  HPanel,
  UiUtil,
  HTB97
//  SoldeCpt,
  , ParamSoc
  , ULIBCLOTURE  // Types et fonctions externes non graphiques
  , uLanceProcess
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
  ;

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
    procedure DeclotureProcessServer ;
    procedure Decloture2Tiers;
    procedure BFermeClick(Sender: TObject);
  private
    Definitive : Boolean ;
    Automatique : Boolean ;
    Exo1 : TExoDate ;
    Exo2 : TExoDate ;
    OnSort : Boolean ;
  public
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  CpteUtil, CpteSav, LicUtil, UtilPgi ;

Function AnnuleclotureComptable(Definitive : Boolean ; FromParam : boolean = false ) : Boolean ;
var FDeClo: TFDEClo;
    OutProg : Boolean ;
    PP : THPanel ;
begin
Result:=FALSE ; if Not _BlocageMonoPoste(True) then Exit ;
if not Definitive and GetParamSocSecur('SO_CPANODYNA', false) and not FromParam then
 begin
  {$IFDEF ENTREPRISE}
  PGIInfo('La gestion des A-nouveaux dynamiques est activée. Nous vous conseillons de désactivez celle ci après ce traitement.') ;
  {$ELSE}
  PGIInfo('La gestion des A-nouveaux dynamiques est activée, désactivez celle ci avant d''annuler la cloture provisoire') ;
  exit ;
  {$ENDIF}
 end ;
FDEClo:=TFDEClo.Create(Application) ; OutProg:=FALSE ;
FDEClo.Definitive:=Definitive ; FDeclo.Automatique:=FALSE ; FDeclo.OnSort:=OutProg ;
if Not Definitive then FDeclo.HelpContext:=7742200 ;
PP:=FindInsidePanel ;
if ((PP=Nil){ or (True)}) then  //SG6 06/12/2004 FQ 14987
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
If OutProg Then Result:=TRUE ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFDECLO.BValiderClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  DeclotureProcessServer ;

  // Attention en cwas, les sonnées étaient rechargés lors du traitement
  //  donc dans le process serveur mais pas au niveau du client !
  ChargeMagExo(False) ;
  GetParamSoc('SO_CPEXOREF', True );
  GetParamSoc('SO_DATECLOTUREPRO', True);

{$ELSE}
  Decloture2Tiers ;
{$ENDIF}

end;

procedure TFDECLO.FormShow(Sender: TObject);
begin
  GBFermeDef.Visible := Definitive ;
  GBFermePro.Visible := Not Definitive ;
  If Definitive
    Then Caption:=Confirmation.Mess[7]
    Else Caption:=Confirmation.Mess[6] ;
  UpdateCaption(Self) ;
  ExoToDates(VH^.EnCours.Code,DateDebN,DateFinN) ;
  If VH^.Precedent.Code<>''
    Then ExoToDates(VH^.Precedent.Code,DateDebM,DateFinM)
    Else
      BEGIN
      DateDebM.Visible:=FALSE ;
      DateFinM.Visible:=FALSE ;
      Label2.Visible:=TRUE ;
      END ;
  If VH^.Suivant.Code<>''
    Then ExoToDates(VH^.Suivant.Code,DateDebN1,DateFinN1)
    Else
      BEGIN
      DateDebN1.Visible:=FALSE ;
      DateFinN1.Visible:=FALSE ;
      Label3.Visible:=TRUE ;
      END
end;

procedure TFDECLO.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  if Parent is THPanel then
    BEGIN
    _DeblocageMonoPoste(True) ;
    Action:=caFree ;
    END ;
end;

procedure TFDECLO.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

procedure TFDECLO.Decloture2Tiers;
var St              : String  ;
    OkDec           : Boolean ;
    i, errID        : Integer ;
    ClotureProcess  : TTraitementCloture ;     // objet qui va faire tout le boulot ;)
begin

  // Confirmation
  If Definitive Then
    BEGIN
    If (not (ctxPCL in V_PGI.PGIContexte)) and ( not V_PGI.FSuperviseur ) Then
      BEGIN
      St:='10;'+Confirmation.Mess[8]+';'+Confirmation.Mess[9]+#13+#10+Confirmation.Mess[10] ;
      Confirmation.Mess[11]:=St ;
      Confirmation.Execute(11,'','') ;
      Exit ;
      END ;
    END ;

  If VH^.Entree.Code<>VH^.EnCours.Code Then
     BEGIN
     Confirmation.Execute(3,'','') ; Exit ;
     END ;

  // Test existance d'écritures sur l'exercice précédent
  // Ajout CA le 18/04/2001
  if Definitive then  // CA - 22/04/2003 - FQ 10193
    begin
    if ((VH^.Precedent.Code <> '') and (not ExisteSQL('SELECT * FROM ECRITURE WHERE E_EXERCICE="'+VH^.Precedent.Code+'"'))) then
      begin
      PGIInfo('Suppression du journal d''à-nouveaux impossible. #10#13L''exercice précédent n''a pas d''écriture.',Caption);
      Exit;
      end;
    end;
  // Fin Ajout CA le 18/04/2001

  // TEST A-Nouveaux lettrable ou pointable // ajout me 12/09/2002
  OkDec := FALSE;
  if (Getparamsoc('SO_CPLIENGAMME') = 'S1') and (ExisteSQL('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" and E_EXERCICE="'+VH^.Encours.Code+'"')) then
    if (PGIAsk('Attention, des à nouveaux lettrable et/ou pointable sont présents, ainsi qu''un lien avec un produit PGI.'+#10#13+' L''export de type synchronisation sera impossible.'+#10#13
              +' Voulez-vous continuez ?','Annulation de clôture'))<> mrYes
      then exit
      else OkDec := TRUE;

  // Placement des exo
  if Definitive then
    BEGIN
    Exo2:=VH^.EnCours ;
    Exo1:=VH^.Precedent ;
    END
  Else
    BEGIN
    Exo2:=VH^.Suivant ;
    Exo1:=VH^.EnCours ;
    END ;

  // Double confirmation
  i:=Confirmation.Execute(0,'','') ;
  If i<>mrYes Then Exit ; Screen.Cursor:=SynCrDefault ;
  i:=Confirmation.Execute(1,'','') ;
  If i<>mrYes Then Exit ; Screen.Cursor:=SynCrDefault ;

  // Mode "patiente"
  HPatienter.Visible:=TRUE ;
  EnableControls(Self,False) ;

  // Création processus
  ClotureProcess := TTraitementCloture.Create( self, Exo1.Code, Exo2.Code, Definitive, Automatique ) ;

  // Traitement
  errID := ClotureProcess.AnnuleCloture( OnSort ) ;

  // Affichage incident
  if ErrId <> CLO_PASERREUR then
    Confirmation.Execute(errID,'','');

  // Libération du processus
  ClotureProcess.Free ;

  // Fin du mode "Patiente"
  EnableControls(Self,TRUE) ;
  HPatienter.Visible:=FALSE ;
  BValider.Visible:=FALSE ;

  // ajout me 12/09/2002
  if OkDec then
    SetParamsoc ('SO_CPDECLOTUREDETAIL',TRUE);

end;

procedure TFDECLO.DeclotureProcessServer;
var St                  : String  ;
    OkDec               : Boolean ;
    i, errID            : Integer ;
    TobParam, TobResult : TOB     ;
    ClotureProcess      : TTraitementCloture ;     // objet qui va faire tout le boulot ;)
    vAnnulClo           : integer;
begin

  // Confirmation
  If Definitive Then
    BEGIN
    If (not (ctxPCL in V_PGI.PGIContexte)) and ( not V_PGI.FSuperviseur ) Then
      BEGIN
      St:='10;'+Confirmation.Mess[8]+';'+Confirmation.Mess[9]+#13+#10+Confirmation.Mess[10] ;
      Confirmation.Mess[11]:=St ;
      Confirmation.Execute(11,'','') ;
      Exit ;
      END ;
    END ;

  If VH^.Entree.Code<>VH^.EnCours.Code Then
     BEGIN
     Confirmation.Execute(3,'','') ; Exit ;
     END ;

  // Test existance d'écritures sur l'exercice précédent
  // Ajout CA le 18/04/2001
  if Definitive then  // CA - 22/04/2003 - FQ 10193
    begin
    if ((VH^.Precedent.Code <> '') and (not ExisteSQL('SELECT * FROM ECRITURE WHERE E_EXERCICE="'+VH^.Precedent.Code+'"'))) then
      begin
      PGIInfo('Suppression du journal d''à-nouveaux impossible. #10#13L''exercice précédent n''a pas d''écriture.',Caption);
      Exit;
      end;
    end;
  // Fin Ajout CA le 18/04/2001

  // TEST A-Nouveaux lettrable ou pointable // ajout me 12/09/2002
  OkDec := FALSE;
  if (Getparamsoc('SO_CPLIENGAMME') = 'S1') and (ExisteSQL('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" and E_EXERCICE="'+VH^.Encours.Code+'"')) then
    if (PGIAsk('Attention, des à nouveaux lettrable et/ou pointable sont présents, ainsi qu''un lien avec un produit PGI.'+#10#13+' L''export de type synchronisation sera impossible.'+#10#13
              +' Voulez-vous continuez ?','Annulation de clôture'))<> mrYes
      then exit
      else OkDec := TRUE;

  // Placement des exo
  if Definitive then
    BEGIN
    Exo2:=VH^.EnCours ;
    Exo1:=VH^.Precedent ;
    END
  Else
    BEGIN
    Exo2:=VH^.Suivant ;
    Exo1:=VH^.EnCours ;
    END ;

  // Double confirmation
  i:=Confirmation.Execute(0,'','') ;
  If i<>mrYes Then Exit ; Screen.Cursor:=SynCrDefault ;
  i:=Confirmation.Execute(1,'','') ;
  If i<>mrYes Then Exit ; Screen.Cursor:=SynCrDefault ;

  // Mode "patiente"
  HPatienter.Visible:=TRUE ;
  EnableControls(Self,False) ;

  // Création processus
  ClotureProcess := TTraitementCloture.Create( self, Exo1.Code, Exo2.Code, Definitive, Automatique ) ;

  // Préparation de la tob
  TobParam := ClotureProcess.CreerTobParam ;
  TobParam.AddChampSupValeur('ONSORT', OnSort) ;

  // Traitement
  TobResult := LanceProcessServer('cgiCloture', 'Decloture', 'aucun', TobParam, True ) ;

  // TobResult bien renseignée ?
  if TobResult.FieldExists('RESULT') then
    begin
    // Récupération du résultat
    errID     := TobResult.GetValue('RESULT') ;
    OnSort    := TobResult.GetValue('ONSORT') ;
    end
  else
    begin
    // Pb avec le process server
    errID     := 13 ;
    OnSort    := False ;
    end ;

  // Libération du processus
  ClotureProcess.Free ;
  TobParam.Free ;
  TobResult.Free ;

  // Fin du mode "Patiente"
  EnableControls(Self,TRUE) ;
  HPatienter.Visible:=FALSE ;
  BValider.Visible:=FALSE ;

  // Affichage incident
  if ErrId <> CLO_PASERREUR then
    Confirmation.Execute(errID,'','');

  // ajout me 12/09/2002
  if OkDec then
    SetParamsoc ('SO_CPDECLOTUREDETAIL',TRUE);

  // --------------------------------------------------------------------------------------
  // ---- On averti le serveur de la mise à jour des paramSoc et de la table exercice  ----
  // --------------------------------------------------------------------------------------
  AvertirCacheServer( 'PARAMSOC' ) ; // SBO : 30/06/2004
  AvertirCacheServer( 'EXERCICE' ) ; // SBO : 30/06/2004

  if (Definitive) and (ErrId = CLO_PASERREUR) then //YMO 11/12/2006 Seulement si le traitement est éxécuté
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
      AvertirCacheServer( 'PARAMSOC' ) ;
    end;
  end;
end;

procedure TFDECLO.BFermeClick(Sender: TObject);
begin
  //SG6 06/12/2004 FQ 14987
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

end.
