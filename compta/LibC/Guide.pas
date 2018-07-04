unit Guide;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, ComCtrls, StdCtrls, ExtCtrls, Spin, Hctrls,
  uTOB,
{$IFDEF EAGLCLIENT}
   UtileAGL,
{$ELSE}
  DBTables,
  MajTable,
  Paramsoc, { AJOUT ME 18/04/2001 }
  dbCtrls,  // Contrôles DB, TNavigateBtn
{$ENDIF}
  {$IFDEF COMPTA}
  Saisie,
  {$ENDIF}
  Hcompte, Buttons, Mask, HEnt1, SaisUtil , GuidUtil, Ent1, hmsgbox,
  GuidTool, Menus, HStatus, Filtre,HSysMenu,
  HTB97, Hqry, HPanel, UiUtil, UtilPGI, SaisComm ;

 Function ParamGuide ( LeQuel : string ; TypeGuide : String3 ; Mode : TActionFiche ): string ;

Const GCXGen=1 ; GCGen=2 ;
      GCXAux=3 ; GCAux=4 ;
      GCXRef=5 ; GCRef=6 ;
      GCXLib=7 ; GCLib=8 ;
      GCXDeb=9 ; GCDeb=10;
      GCXCre=11; GCCre=12;
      GCXMrg=13; GCMrg=14;
      GCXEnc=15; GCTva=16;

type
  TFGuide = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    PEntete: TPanel;
    TJournal: THLabel;
    TDevise: THLabel;
    TNaturePiece: THLabel;
    TEtablissement: THLabel;
    FJournal: THValComboBox;
    FDevise: THValComboBox;
    FNaturePiece: THValComboBox;
    FEtablissement: THValComboBox;
    Messages: THMsgBox;
    PGuide: TPanel;
    TNomGuide: THLabel;
    FNomGuide: THValComboBox;
    POPS: TPopupMenu;
    FTypeContrePartie: THValComboBox;
    FListe: THGrid;
    HG: THCpteEdit;
    HA: THCpteEdit;
    HMrg: THValComboBox;
    Timer1: TTimer;
    HMTrad: THSystemMenu;
    XX_WHERE: TPanel;
    Valide97: TToolbar97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    PPied: TToolbar97;
    BInsLigne: TToolbarButton97;
    BDelLigne: TToolbarButton97;
    BEche: TToolbarButton97;
    BVentil: TToolbarButton97;
    BGuide: TToolbarButton97;
    BAssistant: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    FTva: THValComboBox;
    GuideTreso: TCheckBox;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    Label1: TLabel;
    FDateCreation: TLabel;
    Label2: TLabel;
    FDateModification: TLabel;
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BInsLigneClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FJournalChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BDelLigneClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BVentilClick(Sender: TObject);
    procedure FDeviseChange(Sender: TObject);
    procedure BEcheClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BGuideClick(Sender: TObject);
    procedure BAssistantClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FNaturePieceChange(Sender: TObject);
    procedure FTypeContrePartieChange(Sender: TObject);
    procedure FEtablissementChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FNomGuideClick(Sender: TObject);
    procedure FNomGuideChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject; Ou: Longint;var Cancel: Boolean; Chg: Boolean);
    procedure FListeCellEnter(Sender: TObject; var ACol, ARow: Longint;var Cancel: Boolean);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint;var Cancel: Boolean);
    procedure FListeSetEditText(Sender: TObject; ACol, ARow: Longint;const Value: string);
    procedure Timer1Timer(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure FListeExit(Sender: TObject);
    procedure GuideTresoClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
  private
    Mode                              : TActionFiche ;
    OldCpte,QuelGuide,OldGuide        : String ;
    DEV                               : RDEVISE ;
    FIsModified,GeneResult   : boolean ;
    FTypeTable                        : String ;
    TypeGuide                         : String3 ;
    FSAJAL                            : TSAJAL ;
    Function  CtrlGuide : boolean ;
    Function  RecupRegle ( Gene,Auxi : String ) : String ;
    Function  CtrlEnTete : boolean;
    Function  CtrlAuxiGene( Lig : integer ) : boolean;
    Function  CtrlMontant( Lig : integer ) : boolean;
    Function  CtrlCompta : boolean ;
    Function  CtrlNumeric ( St : string ; Gene : boolean ) : boolean ;
    Function  IsCollectif( Cpte : String ): boolean;
    Function  ChercheNat( Cpte,Table : string ) : String;
    procedure LoadGuide ;
    procedure ClearScreen ;
    procedure ValideLeGuide ;
    procedure AssisteGuide  ;
    function  NomGuideOk : boolean ;
    function  ClickValide : boolean ;
    Procedure GeleBoutonPourConsult ( vBoOn : Boolean = True ) ;
    Procedure SetFoc (C,R : Integer) ;
    procedure VerifGen(Cpte : String ; R : Integer) ;
    procedure TvaSuiv ( Suiv : boolean ) ;
    procedure VerifTva ( ACol,ARow : longint ) ;
    procedure FormateMontant ( ACol,ARow : Longint ) ;
    procedure SetModified( Value : boolean ) ;

    procedure ModeConsult ( vBoOn: boolean = True );

    procedure ChercheGen ; // FQ13580 SBO 04/10/2005 Ajout test comptes interdits / fermés
    procedure ChercheAux ; // FQ13580 SBO 04/10/2005 Ajout test comptes fermés
    function  JalFerme : Boolean ;// FQ13580 SBO 04/10/2005 Ajout test Jal fermés
    function  Bouge (Button : TNavigateBtn) : boolean ;

    property  IsModified : boolean read FIsModified write SetModified ;


 public
  end;


implementation

uses

  GuideAna
  , FichComm
{$IFDEF EAGLCLIENT}
{$ELSE}
  ,PrintDBG
{$ENDIF}
  ;

{$R *.DFM}


Function ParamGuide ( LeQuel : string ; TypeGuide : String3 ; Mode : TActionFiche ): string ;
var FGuide: TFGuide;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],True,'nrSaisieModif') then Exit ;
result:=LeQuel;
FGuide:=TFGuide.Create(Application);
FGuide.Mode:=Mode ;
FGuide.QuelGuide:=LeQuel;
FGuide.TypeGuide:=TypeGuide ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FGuide.ShowModal ;
     result:=FGuide.QuelGuide;
    finally
      FGuide.Free ;
     _Bloqueur('nrSaisieModif',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FGuide,PP) ;
   FGuide.Show ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/11/2002
Modifié le ... :   /  /    
Description .. : - 28/11/2002 - evite le plantage su le echap qd on n'a rine 
Suite ........ : saisie dans les guides
Mots clefs ... : 
*****************************************************************}
procedure TFGuide.FormDestroy(Sender: TObject);
Begin
 FListe.VidePile(true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 27/05/2002
Modifié le ... :   /  /    
Description .. : la case à cocher des guides de trésorerie n'est accessible
Suite ........ : uniquement pour les guides pour les saisies
Mots clefs ... : 
*****************************************************************}
procedure TFGuide.FormShow(Sender: TObject);
begin
VH^.OuiTvaEnc := true ;
{#TVAENC}
if ((Not VH^.OuiTvaEnc) or (TypeGuide='ENC') or (TypeGuide='DEC')) then BEGIN Fliste.ColWidths[GCXEnc]:=0 ; Fliste.ColWidths[GCTva]:=0 ; END ;
GuideTreso.Visible:=TypeGuide='NOR' ;
OldCpte:=''; //BValider.Enabled:=false ;
IsModified:=false ;
if TypeGuide='ABO' then
   BEGIN
   FTypeTable:='ttGuideAbo' ; Caption:=Messages.Mess[22] ;
   END else if ((TypeGuide='ENC') or (TypeGuide='DEC')) then
   BEGIN
   FTypeTable:='ttGuideEncDec' ; Caption:=Messages.Mess[23] ;
   //FNomGuide.Enabled:=False ;
   BInsert.Enabled:=False ; BDelete.Enabled:=False ;
   END else
   BEGIN
   FTypeTable:='ttGuideEcr' ; FJournal.DataType:='ttJalGuide' ; Caption:=Messages.Mess[21] ;
   END ;
FNomGuide.DataType:=FTypeTable ;
if QuelGuide<>'' then BEGIN FNomGuide.Value:=QuelGuide; LoadGuide; END;
if Mode=taConsult then ModeConsult ;
if ((TypeGuide<>'NOR') and (TypeGuide<>'ABO')) then BGuide.Enabled:=False ;

{$IFDEF EAGLCLIENT}
  // On passe par le mul pour sélectionner le guide en eAGL
   FNomGuide.Enabled := True ; //RRO 21032003 (Mode = taCreat) or (Mode = taCreatOne) ; //False ;
{$ELSE}
  { AJOUT ME 18/04/2001 }
  //RR préparation version S3 Premium.
  if (((ctxPCL in V_PGI.PGIContexte) or (EstComptaSansAna)) and (GetParamsoc('SO_CPPCLSANSANA')=True)) then
    BVentil.Visible:= False;
  { FIN AJOUT ME 18/04/2001 }
{$ENDIF}

{$IFDEF COMPTA}
{$ELSE}
  BGuide.Visible := False;
{$ENDIF}
UpdateCaption(Self) ;
  {JP 28/04/05 : FQ 15784 : on cache le bouton tant que la table ANAGUI n'aura pas été mise à jour}
  BVentil.Visible := not VH^.AnaCroisaxe;

end;

procedure TFGuide.BInsLigneClick(Sender: TObject);
begin
FListe.InsertRow(FListe.Row) ;
G_Renum(FListe) ;
FListe.SetFocus ;
end;

Procedure TFGuide.SetFoc (C,R : Integer) ;
BEGIN
FListe.Col:=C ; FListe.Row:=R ; FListe.SetFocus ;
END ;

Function TFGuide.CtrlGuide : boolean;
var i    : integer;
    Good : boolean;

    function TotalementVide(Ligne : Integer ) : boolean ;
    var i : integer ;
    begin
      Result := False;
      for i:=1 to FListe.ColCount-1 do
        if FListe.Cells[i,Ligne]<>'' then Exit;
      Result := True;
    end;
BEGIN
  CtrlGuide:=False ;
  // contrôles sur l'entête du guide
  if Not CtrlEnTete then Exit ;
  // contrôles sur les lignes
  For i:=1 to FListe.RowCount-1 do BEGIN
    if Not G_LigneVide(FListe,i) then BEGIN
      Good:=CtrlAuxiGene(i) ;
      if Good then Good:=CtrlMontant(i);
      if Good then Good:=CtrlFormule(FListe,i,Messages,FALSE);
      if Not Good then BEGIN FListe.Row:=i ; exit; END ;
      END
    else BEGIN
      if IsModified then BEGIN
        SetFoc(GCGen,i) ;
        if FListe.RowCount=2 then BEGIN
          Messages.Execute(28,'','') ; Exit ;
          END
        else BEGIN
          if not TotalementVide(i) then begin // FQ 10207
            Messages.Execute(25,'','');
            Exit;
          end;
        END;
      END;
    END;
  END;
  Good:=CtrlCompta;
  if Not Good then exit;
  CtrlGuide:=True;
END;

Function TFGuide.CtrlEnTete : boolean;
BEGIN
CtrlEnTete:=False ;
if Not NomGuideOK then BEGIN FNomGuide.SetFocus ; Exit ; END;
if FDevise.Value='' then BEGIN Messages.Execute(35,'',''); FDevise.SetFocus ; Exit ; END;
{ MODIF ME 18/04/2001 }
//if ((not (ctxStandard in V_PGI.PGIContexte)) and (FEtablissement.Value='')) then BEGIN Messages.Execute(40,'','') ; FEtablissement.SetFocus ; Exit ; END;
if FJournal.Value<>'' then if ((FJournal.Value=VH^.JalATP) or (FJournal.Value=VH^.JalVTP)) then BEGIN Messages.Execute(46,'','') ; FJournal.SetFocus ; Exit ; END ;   
CtrlEnTete:=True ;
END ;

Function TFGuide.IsCollectif(Cpte:String):boolean;
var Q      : TQuery;
    SQL    : string;
BEGIN
IsCollectif:=True;
SQL:='Select G_COLLECTIF from GENERAUX where G_GENERAL="'+Cpte+'"';
Q:=OpenSQL(SQL,true);
if Not Q.EOF then
   BEGIN
   if Q.FindField('G_COLLECTIF').AsString='X' then BEGIN Ferme(Q) ; Exit ; END ;
   END ;
Ferme(Q);
IsCollectif:=False;
END;

Function TFGuide.RecupRegle ( Gene,Auxi : String ) : String ;
Var Q : TQuery ;
    Nat : String ;
BEGIN
Result:='' ;
if Auxi<>'' then
   BEGIN
   Q:=OpenSQL('Select T_MODEREGLE from TIERS Where T_AUXILIAIRE="'+Auxi+'"',True) ;
   if Not Q.EOF then Result:=Q.Fields[0].AsString ;
   Ferme(Q) ;
   END else
   BEGIN
   Q:=OpenSQL('Select G_MODEREGLE, G_NATUREGENE from GENERAUX Where G_GENERAL="'+Gene+'"',True) ;
   if Not Q.EOF then
      BEGIN
      Nat:=Q.Fields[1].AsString ;
      if ((Nat='TID') or (Nat='TIC')) then Result:=Q.Fields[0].AsString ;
      END ;
   Ferme(Q) ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/11/2002
Modifié le ... : 06/07/2004
Description .. : - 28/11/2002 - fiche 10482 - plus de controle sur les TIC ou 
Suite ........ : TID
Suite ........ : - LG - FB 13773 - 06/07/2004 - on pouvait saisir des 
Suite ........ : racines de compte
Mots clefs ... : 
*****************************************************************}
Function TFGuide.CtrlAuxiGene(Lig:integer) : boolean;
var CBG,CBA                        : Boolean ;
    GExist,AExist                  : boolean;
    Collectif,GClose,AClose        : boolean ;
    ModeR,NatGene,NatAuxi,NatPiece : String ;
BEGIN
CtrlAuxiGene:=False ; ModeR:='' ;
HG.text:=FListe.Cells[GCGen,Lig] ; CBG:=G_Croix(FListe.Cells[GCXGen,Lig]);
HA.text:=FListe.Cells[GCAux,Lig] ; CBA:=G_Croix(FListe.Cells[GCXAux,Lig]);
GExist:=(HG.ExisteH>0) ; if GExist and not CBG then FListe.Cells[GCGen,Lig] := HG.text ;
GClose:=IsClose(HG.Text,'GENERAUX') ;
AExist:=(HA.ExisteH>0) ; if AExist and not CBA then FListe.Cells[GCAux,Lig] := HA.text ;
AClose:=IsClose(HA.Text,'TIERS') ;
Collectif:=IsCollectif(HG.Text) ;  NatPiece:=FNaturePiece.Value ;
if GExist then NatGene:=ChercheNat(HG.Text,'GENERAUX') else NatGene:='' ;
if AExist then NatAuxi:=ChercheNat(HA.Text,'TIERS') else NatAuxi:='' ; 
{#TVAENC}
if VH^.OuiTvaEnc then
   BEGIN
   if ((NatGene='') or ((NatGene<>'CHA') and (NatGene<>'PRO') and (NatGene<>'IMO'))) then
      BEGIN
      FListe.Cells[GCXEnc,Lig]:='' ; FListe.Cells[GCTva,Lig]:='' ;
      END ;
   END ;

//// ====== CONTROLES POUR TOUS TYPES DE GUIDE ====== ////

// Compte interdit ? // FQ13580 SBO 04/10/2005 Ajout test comptes interdits
if EstInterdit( FSAJAL.COMPTEINTERDIT, HG.Text, 0) > 0 then
  begin
  SetFoc(GCGen,Lig) ;
  PgiBox('Ce compte général est interdit sur ce journal.', Caption);
  Exit;
  end ;

if (GExist and GClose) then BEGIN SetFoc(GCGen,Lig) ; Messages.Execute(0,'',''); Exit; END;
if (AExist and AClose) then BEGIN SetFoc(GCAux,Lig) ; Messages.Execute(1,'',''); Exit; END;
//ctrl de cohérence entre nature de pièce et nature de cpte général
if NatGene<>'' then
   BEGIN
   if ((NatGene='COC') and ((NatPiece='FF') or (NatPiece='AF') or (NatPiece='RF') or (NatPiece='OF'))) then
       BEGIN SetFoc(GCGen,Lig) ; Messages.Execute(42,'','') ; Exit ; END ;
   if ((NatGene='COF') and ((NatPiece='FC') or (NatPiece='AC') or (NatPiece='RC') or (NatPiece='OC'))) then
       BEGIN SetFoc(GCGen,Lig) ; Messages.Execute(42,'','') ; Exit ; END ;
//   if ((NatGene='TID') and ((NatPiece='FF') or (NatPiece='AF') or (NatPiece='RF') or (NatPiece='OF'))) then
//       BEGIN SetFoc(GCGen,Lig) ; Messages.Execute(42,'','') ; Exit ; END ;
//   if ((NatGene='TIC') and ((NatPiece='FC') or (NatPiece='AC') or (NatPiece='RC') or (NatPiece='OC'))) then
//       BEGIN SetFoc(GCGen,Lig) ; Messages.Execute(42,'','') ; Exit ; END ;
   END ;
//ctrl de cohérence entre nature de général et nature d'auxiliaire
if ((NatGene<>'') and (NatAuxi<>'')) then
   BEGIN
   if ((NatGene='COC') and ((NatAuxi='AUC') or (NatAuxi='FOU') or (NatAuxi='SAL'))) then
       BEGIN SetFoc(GCAux,Lig) ; Messages.Execute(43,'','') ; Exit ; END ;
   if ((NatGene='COF') and ((NatAuxi='AUD') or (NatAuxi='CLI') or (NatAuxi='SAL'))) then
       BEGIN SetFoc(GCAux,Lig) ; Messages.Execute(43,'','') ; Exit ; END ;
   if ((NatGene='COS') and (NatAuxi<>'SAL') and (NatAuxi<>'DIV')) then
       BEGIN SetFoc(GCAux,Lig) ; Messages.Execute(43,'','') ; Exit ; END ;
   END ;

//// ====== CONTROLES EN FONCTION DU TYPE DE GUIDE ====== ////
if ((TypeGuide<>'ENC') and (TypeGuide<>'DEC')) then
   BEGIN
   if FListe.Cells[GCMrg,Lig]='' then
      BEGIN
      ModeR:=RecupRegle(HG.Text,HA.Text) ;
      HMrg.Value:=ModeR ;
      if ModeR<>'' then FListe.Cells[GCMrg,Lig]:=HMrg.Text ;
      END ;
   END ;
if ((TypeGuide='ENC') or (TypeGuide='DEC')) then
   BEGIN
   if ((Lig=1) and (Not GExist)) then BEGIN SetFoc(GCGen,Lig); Messages.Execute(27,'',''); Exit; END;
   END else
if TypeGuide='ABO' then
   BEGIN
   if Not GExist then BEGIN SetFoc(GCGen,Lig); Messages.Execute(27,'',''); Exit; END;
   if ((Collectif) and (Not AExist)) then BEGIN SetFoc(GCAux,Lig); Messages.Execute(27,'',''); Exit; END;
   END else
if ((TypeGuide='POI') or (TypeGuide='NOR')) then
   BEGIN
   if GExist and Collectif and (HA.text='') and Not CBA then FListe.Cells[GCXAux,Lig]:='X' ;
   if (HG.Text<>'') then
      if ((Pos('*',HG.text)>0) or (Pos('?',HG.Text)>0) or Not GExist) and Not CBG then
         BEGIN SetFoc(GCGen,Lig) ; Messages.Execute(2,'',''); Exit; END;
   if (HA.Text<>'') then
      if ((Pos('*',HA.text)>0) or (Pos('?',HA.Text)>0) or Not AExist) and Not CBA then
         BEGIN SetFoc(GCAux,Lig); Messages.Execute(3,'',''); Exit; END;
   if (HG.text='') and (HA.Text='') and Not CBG and Not CBA then
       BEGIN SetFoc(GCGen,Lig); Messages.Execute(4,'',''); Exit; END;
   if (HA.text<>'') or CBA then
       if GExist and Not Collectif then BEGIN SetFoc(GCGen,Lig); Messages.Execute(5,'',''); Exit; END;
   if ((HG.Text<>'') and (Not CtrlNumeric(HG.Text,True))) then
       BEGIN SetFoc(GCGen,Lig); Messages.Execute(9,'',''); Exit; END;
   END;
CtrlAuxiGene:=True;
END;

Function TFGuide.CtrlNumeric ( St : string ; Gene : boolean ): boolean ;
Var Good : boolean;
BEGIN
Good:=IsMontant(St,Gene) ;
if Not Good then Good:=(Pos('[',St)>0);   // si pas numérique, peut être une formule
CtrlNumeric:=Good;
END;

Function TFGuide.CtrlMontant(Lig:integer) : boolean;
var Deb, Cred               : String ;
    CBDeb,CBCred            : Boolean ;
BEGIN
CtrlMontant:=False;
Deb:=FListe.Cells[GCDeb,Lig] ; Cred:=FListe.Cells[GCCre,Lig] ;
CBDeb:=G_Croix(FListe.Cells[GCXDeb,Lig]) ;
CBCred:=G_Croix(FListe.Cells[GCXCre,Lig]) ;

//// ====== CONTROLES POUR TOUS TYPES DE GUIDE ====== ////
if (Deb<>'') and (Cred<>'') then
    BEGIN SetFoc(GCDeb,Lig); Messages.Execute(7,'',''); Exit; END;
if ((DEB<>'') and (Not CtrlNumeric(Deb,False))) or
   ((CRED<>'') and (Not CtrlNumeric(Cred,False))) then
    BEGIN SetFoc(GCDeb,Lig); Messages.Execute(10,'',''); Exit; END;
if CtrlNumeric(Deb,False) and (Pos('[',Deb)<=0) and (Pos('-',Deb)>0) and Not VH^.MontantNegatif then
    BEGIN SetFoc(GCDeb,Lig); Messages.Execute(11,'',''); Exit; END;
if CtrlNumeric(Cred,False) and (Pos('[',Cred)<=0) and (Pos('-',Cred)>0)and Not VH^.MontantNegatif then
    BEGIN SetFoc(GCCre,Lig); Messages.Execute(11,'',''); Exit; END;
if ((Valeur(DEB)<0) and (Pos('[',Deb)<=0) and (Not VH^.MontantNegatif)) then
    BEGIN SetFoc(GCDeb,Lig); Messages.Execute(38,'',''); Exit; END;
if ((Valeur(CRED)<0) and (Pos('[',Cred)<=0) and (Not VH^.MontantNegatif)) then
    BEGIN SetFoc(GCCre,Lig); Messages.Execute(38,'',''); Exit; END;
//// ====== CONTROLES EN FONCTION DU TYPE DE GUIDE ====== ////
if TypeGuide='ABO' then
   BEGIN
   if ((Deb='') and (Cred='')) or ((Valeur(DEB)=0) and (Valeur(CRED)=0)) then
       BEGIN SetFoc(GCDeb,Lig); Messages.Execute(26,'',''); Exit; END;
   if ((Deb<>'') and (Not IsMontant(Deb,False))) then
       BEGIN SetFoc(GCDeb,Lig); Messages.Execute(26,'',''); Exit; END;
   if ((Cred<>'') and (Not IsMontant(Cred,False))) then
       BEGIN SetFoc(GCCre,Lig); Messages.Execute(26,'',''); Exit; END;
   END else
if ((TypeGuide='ENC') or (typeGuide='DEC')) then
   BEGIN
   // les montants peuvent être non renseignés et sans arrêt
   END else
if ((TypeGuide='POI') or (TypeGuide='NOR')) then
   BEGIN
   if (Deb='') and (Cred='') and Not CBDeb and Not CBCred then
       BEGIN SetFoc(GCDeb,Lig); Messages.Execute(17,'',''); Exit; END;
   END;
(* GG A priori, rien n'empèche de s'arrèter sur un montant ??? A valider quand même (17/09/96)
if (Deb.Text<>'') and (CBDeb.Checked or CBCred.Checked) then
    BEGIN Deb.SetFocus; SB.HorzScrollBar.Position:=PosHSB; Messages.Execute(8,'',''); Exit; END;
if (Cred.Text<>'') and (CBDeb.Checked or CBCred.Checked) then
    BEGIN Cred.SetFocus; SB.HorzScrollBar.Position:=PosHSB; Messages.Execute(8,'',''); Exit; END;
*)
CtrlMontant:=True;
END;

Function TFGuide.CtrlCompta : boolean;
var Q      : TQuery;
    SQL,St : string;
    OkContrePartie,MulDev,Equil : boolean;
    C,D    : double;
    i      : integer;
    ContrePartieJal : String17 ;
BEGIN
CtrlCompta:=False; MulDev:=False;
//// ====== CONTROLES EN FONCTION DU TYPE DE GUIDE ====== ////
if ((TypeGuide='ABO') or (TypeGuide='ENC') or (TypeGuide='DEC')) then
   BEGIN
   if FJournal.Text='' then
      BEGIN
      FJournal.SetFocus ; Messages.Execute(29,'','') ; Exit ;
      END;
   if FNaturePiece.Text='' then
      BEGIN
      FNaturePiece.SetFocus ; Messages.Execute(30,'','') ; Exit ;
      END;
   END;
if ((TypeGuide='ENC') or (TypeGuide='DEC')) then
   BEGIN
   if FDevise.Text='' then
      BEGIN
      FDevise.SetFocus ; Messages.Execute(31,'','') ; Exit ;
      END;
   END;
if TypeGuide='ABO' then
   BEGIN
   C:=0; D:=0;
   For i:=1 to FListe.RowCount-1 do
      BEGIN
      C:=C+Valeur(FListe.Cells[GCCre,i]);
      D:=D+Valeur(FListe.Cells[GCDeb,i]);
      END;
   Equil:=(Arrondi(D-C,V_PGI.OkDecE)=0) ;
   if Not Equil then BEGIN Messages.Execute(32,'',''); exit; END;
   END;

//// ====== CONTROLES POUR TOUS TYPES DE GUIDE ====== ////
if (FDevise.text<>'') and (FDevise.value<>V_PGI.DevisePivot) then
   BEGIN
   SQL:='Select J_MULTIDEVISE from JOURNAL where J_JOURNAL="'+FJournal.Value+'"';
   Q:=OpenSQL(SQL,true);
   if Not Q.EOF then MulDev:=(Q.FindField('J_MULTIDEVISE').AsString='X');
   Ferme(Q);
   if Not MulDev then
      BEGIN
      FJournal.SetFocus ; Messages.Execute(14,'','') ; Exit ;
      END;
   END;
if FJournal.Value<>'' then
   BEGIN
   Q:=OpenSQL('Select J_COMPTEURNORMAL, J_COMPTEURSIMUL,J_NATUREJAL,J_CONTREPARTIE from JOURNAL Where J_JOURNAL="'+FJournal.Value+'"',True) ;
   if Not Q.EOF then
      BEGIN
      if ((Q.Fields[0].AsString='') or (Q.Fields[1].AsString='')) then BEGIN Ferme(Q) ; Messages.Execute(37,'','') ; Exit ; END ;
      // CP 30/09/97 Contrôle sur compte de contrepartie du journal de banque / Caisse
      if (Q.Fields[2].AsString='BQE') or (Q.Fields[2].AsString='CAI') then
        BEGIN
        St:=Q.Fields[3].AsString ;
        ContrePartieJal:=TrouveAuto(St,1) ;
        OkContrePartie:=False ;
        for i:=1 to FListe.RowCount-1 do
          BEGIN
          if (Fliste.Cells[GCGen,i]=ContrePartieJal) then BEGIN OkContrePartie:=True ; Break ; END ;
          END ;
        if Not OkContrePartie then BEGIN Ferme(Q) ; Messages.Execute(44,'','') ; Exit ; END ;
        END ;
      END ;
   Ferme(Q) ;
   END else
   BEGIN
   Messages.Execute(39,'','') ; Exit ;
   END ;
{ MODIF ME LE 18/04/2001 }
//if ((not (ctxStandard in V_PGI.PGIContexte)) and (FEtablissement.ItemIndex=-1)) then BEGIN Messages.Execute(40,'','') ; Exit ; END ;
CtrlCompta:=True;
END;

function TFGuide.ClickValide : boolean ;
BEGIN
ClickValide:=False ;
if JalFerme then Exit ;
if Not CtrlGuide then BEGIN IsModified:=True ; GeneResult:=False ; Exit ; END ;
if Transactions(ValideLeGuide,5)<>oeOk then BEGIN MessageAlerte(Messages.Mess[34]) ; Exit ; END ;
ClickValide:=True ;
END ;

procedure TFGuide.BValiderClick(Sender: TObject);
begin ClickValide ; end ;

function TFGuide.NomGuideOk : boolean ;
var Q : TQuery ;
BEGIN
Result:=true ;
if FNomGuide.Text='' then BEGIN Messages.Execute(20,'','') ; Result:=false ; Exit ; END;
Q:=OpenSQL('Select GU_GUIDE From GUIDE Where GU_TYPE="'+TypeGuide+'" AND GU_LIBELLE="'+FNomGuide.Text+'"', True) ;
While Not Q.EOF do
   BEGIN
   if ((FNomGuide.ItemIndex<0) or ((FNomGuide.ItemIndex>=0) and (Q.Fields[0].AsString<>FNomGuide.Value))) then
      BEGIN
      Messages.Execute(41,'','') ; Result:=False ; Break ;
      END ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/05/2002
Modifié le ... :   /  /    
Description .. : suppression de warning
Mots clefs ... : 
*****************************************************************}
procedure TFGuide.ValideLeGuide ;
Var nb,i,ia,ix,ind : integer ;
    NomG,St,St1,ValSt  : String ;
    T                  : TVentGuide ;
    Q                  : TQuery ;
    New                : boolean ;
    TGuide             : TOB ;
    TEcrGuide          : TOB ;
    TAnaGuide          : TOB ;
begin
  GeneResult := False ;
  NomG       := FNomGuide.Text ;
  if FNomGuide.ItemIndex<0 then
    BEGIN
    New:=True;
    Q:=OpenSQL('Select MAX(GU_GUIDE) from GUIDE Where GU_TYPE="'+TypeGuide+'" Order by 1',True);// Order by GU_GUIDE
    if not Q.EOF then
      begin
      ValSt := Q.Fields[0].AsString ;
      if ValSt<>'' then
        begin
        i     := StrToInt(ValSt);
        ValSt := IntToStr(i+1);
        end //ItIndex:=i; END
      else
        ValSt:='001' ;
      while Length(ValSt)<3 do
        ValSt:='0'+ValSt;
      end
    else
      ValSt := '001' ;
    FNomGuide.Value := ValSt;
    Ferme(Q) ;
    end
  else
    begin
    ValSt := FNomGuide.Value;
    New   := false;
    end;

  // On supprime pour ré-insérer
  ExecuteSQL('DELETE FROM GUIDE WHERE GU_TYPE="'+TypeGuide+'" AND GU_GUIDE="'+ValSt+'"') ;
  ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="'+TypeGuide+'" AND EG_GUIDE="'+ValSt+'"') ;
  ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="'+TypeGuide+'" AND AG_GUIDE="'+ValSt+'"') ;
  nb := 0 ;

  // ---------------
  // DEBUT MAJ GUIDE
  TGuide := TOB.Create('GUIDE',nil,-1);
  TGuide.InitValeurs ;
  TGuide.PutValue('GU_TYPE',        TypeGuide) ;
  TGuide.PutValue('GU_GUIDE',       ValSt) ;
  TGuide.PutValue('GU_LIBELLE',     NomG) ;
  TGuide.PutValue('GU_TRESORERIE',  CheckToString(GuideTreso.Checked)) ;
  if New then
    TGuide.PutValue('GU_DATECREATION', Date)
  else
    begin
    if Trim(FDateCreation.Caption) = '' then
      FDateCreation.Caption := DateTimeToStr(Date) ;
    if Trim(FDateModification.Caption) = '' then
      FDateModification.Caption := DateTimeToStr(Date) ;
     TGuide.PutValue('GU_DATECREATION', StrToDateTime(FDateCreation.Caption)) ;
    end ;
  TGuide.PutValue('GU_DATEMODIF',      Date) ;
  TGuide.PutValue('GU_JOURNAL',        FJournal.Value) ;
  TGuide.PutValue('GU_NATUREPIECE',    FNaturePiece.Value) ;
  TGuide.PutValue('GU_ETABLISSEMENT',  FEtablissement.Value) ;
  TGuide.PutValue('GU_TYPECTREPARTIE', FTypeContrePartie.Value) ;
  TGuide.PutValue('GU_DEVISE',         FDevise.Value) ;
  TGuide.PutValue('GU_TYPE',           TypeGuide) ;
  TGuide.PutValue('GU_UTILISATEUR',    V_PGI.User) ;
  TGuide.PutValue('GU_SOCIETE',        V_PGI.CodeSociete) ;
  // insertion base
  TGuide.InsertDB(nil) ;
  TGuide.Free ;
  // FIN MAJ GUIDE
  // -------------

  // ============================================================
  // DETAIL DU GUIDE (lignes d'écritures générales et analytiques)
  // ============================================================

  InitMove(FListe.RowCount,'') ;
  for i:=1 to FListe.RowCount-1 do
    if Not G_LigneVide(FListe,i) then
      begin
      Inc(nb) ;
      St := '' ;
      MoveCur(FALSE) ;
      // ------------------
      // DEBUT MAJ ECRGUI
      TEcrGuide := TOB.Create('ECRGUI',nil,-1) ;
      TEcrGuide.InitValeurs ;
      TEcrGuide.PutValue('EG_TYPE',     TypeGuide ) ;
      TEcrGuide.PutValue('EG_GUIDE',    ValSt ) ;
      TEcrGuide.PutValue('EG_NUMLIGNE', nb ) ;
      // FQ 12090
      if (FListe.Cells[GCAux,i] = '') then TEcrGuide.PutValue('EG_RIB',      '' )
      else begin
        Q := OpenSQL('SELECT R_CODEIBAN, R_PAYS, R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB, R_DOMICILIATION FROM RIB WHERE R_AUXILIAIRE="'+FListe.Cells[GCAux,i]+'" AND R_PRINCIPAL="X"', True);
        if Q.EOF then TEcrGuide.PutValue('EG_RIB', '')
        else begin
          if (Q.FindField('R_CODEIBAN').AsString = '') then TEcrGuide.PutValue('EG_RIB', calcRIB(Q.FindField('R_PAYS').AsString, Q.FindField('R_ETABBQ').AsString, Q.FindField('R_GUICHET').AsString, Q.FindField('R_NUMEROCOMPTE').AsString, Q.FindField('R_CLERIB').AsString))
                                                       else TEcrGuide.PutValue('EG_RIB', Q.FindField('R_CODEIBAN').AsString);
        end;
      end;
      TEcrGuide.PutValue('EG_GENERAL',  FListe.Cells[GCGen,i]) ;
      if G_Croix(FListe.Cells[GCXGen,i])
        then St := St + 'X'
        else St := St + '-' ;
      TEcrGuide.PutValue('EG_AUXILIAIRE', FListe.Cells[GCAux,i]) ;
      if G_Croix(FListe.Cells[GCXAux,i])
        then St := St + 'X'
        else St := St + '-' ;
      TEcrGuide.PutValue('EG_REFINTERNE', FListe.Cells[GCRef,i] ) ;
      if G_Croix(FListe.Cells[GCXRef,i])
        then St := St + 'X'
        else St := St + '-' ;
      TEcrGuide.PutValue('EG_LIBELLE', FListe.Cells[GCLib,i] ) ;
      if G_Croix(FListe.Cells[GCXLib,i])
        then St := St + 'X'
        else St := St + '-' ;
      TEcrGuide.PutValue('EG_DEBITDEV', FListe.Cells[GCDeb,i] ) ;
      if G_Croix(FListe.Cells[GCXDeb,i])
        then St := St + 'X'
        else St := St + '-' ;
      TEcrGuide.PutValue('EG_CREDITDEV', FListe.Cells[GCCre,i] ) ;
      if G_Croix(FListe.Cells[GCXCre,i])
        then St := St + 'X'
        else St := St + '-' ;
      HMrg.Libelle := FListe.Cells[GCMrg,i] ;
      TEcrGuide.PutValue('EG_MODEREGLE', HMrg.Value) ;
      if G_Croix(FListe.Cells[GCXMrg,i])
        then St := St + 'X'
        else St := St + '-' ;
      TEcrGuide.PutValue('EG_ARRET', St) ;
      {#TVAENC}
      if VH^.OuiTvaEnc then
        begin
        St := FListe.Cells[GCXEnc,i] ;
        if ((St<>'X') and (St<>'-'))
          then St:='' ;
        TEcrGuide.PutValue('EG_TVAENCAIS', St) ;
        St  := FListe.Cells[GCTva,i] ;
        Ind := FTva.Items.IndexOf(St) ;
        if Ind>=0
          then st := FTva.Values[Ind]
          else st := '' ;
        TEcrGuide.PutValue('EG_TVA', St ) ;
        end ;
      // Insertion base
      TEcrGuide.InsertDB(nil) ;
      TEcrGuide.Free ;
      // FIN DE MAJ DE ECRGUI
      // --------------------

      // ----------------
      // DEBUT MAJ ANAGUI
      T:=TVentGuide(FListe.Objects[0,i]) ;
      if T<>NIL then
      for ix:=1 to MaxAxe do
        for ia:=0 to T.Ventil[ix].Count-1 do
          if Trim(T.Ventil[ix].Strings[ia])<>'' then
            begin
            St1:=T.Ventil[ix].Strings[ia] ;
            St:='' ;
            TAnaGuide := TOB.Create('ANAGUI',nil,-1) ;
            TAnaguide.InitValeurs ;
            TAnaGuide.PutValue('AG_TYPE',       TypeGuide) ;
            TAnaGuide.PutValue('AG_GUIDE',      ValSt) ;
            TAnaGuide.PutValue('AG_NUMLIGNE',   nb) ;
            TAnaGuide.PutValue('AG_NUMVENTIL',  ia+1) ;
            TAnaGuide.PutValue('AG_AXE',        'A'+IntToStr(ix)) ;
            TAnaGuide.PutValue('AG_SECTION',    Trim(Copy(St1,1,35))) ;
              Delete(St1,1,35) ;
              St:=St+St1[1] ;
              Delete(St1,1,1) ;
            TAnaGuide.PutValue('AG_POURCENTAGE',Trim(Copy(St1,1,100)) ) ;
              Delete(St1,1,100) ;
              St:=St+St1[1] ;
              Delete(St1,1,1) ;
            TAnaGuide.PutValue('AG_POURCENTQTE1', Trim(Copy(St1,1,100)) ) ;
              Delete(St1,1,100) ;
              St:=St+St1[1] ;
              Delete(St1,1,1) ;
            TAnaGuide.PutValue('AG_POURCENTQTE2', Trim(Copy(St1,1,100)) ) ;
              Delete(St1,1,100) ;
              St:=St+St1[1] ;
              Delete(St1,1,1) ;
            TAnaGuide.PutValue('AG_ARRET', St ) ;
            // insertion base
            TAnaGuide.InsertDB(nil) ;
            TAnaGuide.Free ;
            end ;
      // --------------
      // FIN MAJ ANAGUI
    end ;

  FiniMove ;

  {$IFDEF AGL525}
    ClearLaTable(FNomGuide.DataType) ;
    FNomGuide.Reload ;    // pour compatibilité AGL < 530
  {$ELSE}
    VideLaTablette(FNomGuide.DataType) ; // AGL 530 ou +
  {$ENDIF}

  FNomGuide.Reload ;
  if TypeGuide<>'ABO' then
    begin
    FDateModification.Caption:=DateTimeToStr(Date) ;
    if New then
      FDateCreation.Caption:=DateTimeToStr(Date);
    GeneResult:=True ;
    end;

 // BValider.Enabled  := false;
  IsModified        := false ;
  FNomGuide.Value  := ValSt ;
  QuelGuide         := ValSt;
 // FNomGuide.Enabled := False ;
end;

procedure TFGuide.FJournalChange(Sender: TObject);
Var Q : TQuery ;
    ModeS : String ;
begin
if TypeGuide<>'ABO' then
   BEGIN
   FDEVISE.Enabled:=True ;
   END else
   BEGIN
   FDEVISE.Value:=V_PGI.DevisePivot ;
   FDEVISE.Enabled:=False ;
   END ;

Q:=OpenSQL('Select J_NATUREJAL, J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="'+FJournal.Value+'"',TRUE) ;
if Not Q.EOF then
   BEGIN
   Case CaseNatJal(Q.Fields[0].AsString) of
      tzJVente  : BEGIN FNaturePiece.DataType:='ttNatPieceVente'  ; FNaturePiece.Value:='FC' ; END ;
      tzJAchat  : BEGIN FNaturePiece.DataType:='ttNatPieceAchat'  ; FNaturePiece.Value:='FF' ; END ;
      tzJBanque : BEGIN FNaturePiece.DataType:='ttNatPieceBanque' ; FNaturePiece.Value:='RC' ; END ;
      tzJOD     : BEGIN FNaturePiece.DataType:='ttNaturePiece'    ; FNaturePiece.Value:='OD' ; END ;
      END ;
   ModeS:=Q.Fields[1].AsString ;
   if ((ModeS<>'') and (ModeS<>'-')) then BGuide.Enabled:=False else
     if ((TypeGuide='NOR') or (TypeGuide='ABO')) then BGuide.Enabled:=True ;

   // Chargement du journal // FQ13580 SBO 04/10/2005
   FSAJAL.Free ;
   if ( FJournal.Value ) <> '' then
     FSAJAL    := TSAJAL.Create(FJournal.Value,False) ;
   END ;
Ferme(Q) ;
IsModified:=true;
end;

procedure TFGuide.BFermeClick(Sender: TObject);
begin
if Not IsInside(Self) then Close ;
end;

procedure TFGuide.BDelLigneClick(Sender: TObject);
begin
if ((Fliste.Row<=0) or (FListe.RowCount<=2)) then Exit ;
if FListe.Objects[0,FListe.Row]<>NIL then TVentGuide(FListe.Objects[0,FListe.Row]).Free ;
FListe.DeleteRow(FListe.Row) ;
G_Renum(FListe) ;
HMrg.Visible:=False ;
FListe.SetFocus ;
end;

procedure TFGuide.VerifTva ( ACol,ARow : longint ) ;
Var St : String ;
    Ind : integer ;
BEGIN
St:=FListe.Cells[ACol,ARow] ;
if St='' then Exit ;
Ind:=FTva.Items.IndexOf(St) ;
if Ind<0 then FListe.Cells[ACol,ARow]:='' ;
END ;

procedure TFGuide.TvaSuiv ( Suiv : boolean ) ;
Var St  : String ;
    Ind : integer ;
BEGIN
if FListe.Col<>GCTva then Exit ;
if FTva.Values.Count<=0 then Exit ;
St:=FListe.Cells[FListe.Col,FListe.Row] ;
if St='' then
   BEGIN
   if Suiv then Ind:=0 else Ind:=FTva.Items.Count-1 ;
   END else
   BEGIN
   FTva.Libelle:=St ; Ind:=FTva.ItemIndex ;
   if Suiv then
      BEGIN
      if Ind=FTva.Values.Count-1 then Ind:=0 else Inc(Ind) ;
      END else
      BEGIN
      if Ind>0 then Dec(Ind) else Ind:=FTva.Values.Count-1 ;
      END ;
   END ;
FTva.ItemIndex:=Ind ;
FListe.Cells[FListe.Col,FListe.Row]:=FTva.Text ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/05/2002
Modifié le ... : 29/11/2002
Description .. : - 24/05/2002 - Gestion du VK_ESCAPE
Suite ........ : - 29/11/2002 - fiche bug 10773 - suppression de 
Suite ........ : VK_ESCAPE -> bug si la fiche etait vide
Mots clefs ... : 
*****************************************************************}
procedure TFGuide.FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ;
Case Key of
    VK_F3  : if ((Vide) and (Screen.ActiveControl=FListe)) then BEGIN Key:=0 ; TvaSuiv(False) ; END ;
    VK_F4  : if ((Vide) and (Screen.ActiveControl=FListe)) then BEGIN Key:=0 ; TvaSuiv(True) ; END ;
    VK_F5  : if Vide then
               if (Screen.ActiveControl=FListe)
                 then begin
                      Key:=0 ;
                      case FListe.Col of
                           GCGen : ChercheGen ;
                           GCAux : ChercheAux ;
                           else    FListeDblClick( nil )
                           end ;
                      end
                 else begin
                      Key:=0 ;
                      AssisteGuide ;
                      end ;
    VK_F10 : if Vide then BEGIN Key:=0 ; ClickValide ; END ;
 VK_RETURN : if ((Screen.ActiveControl=Fliste) and (Vide)) then KEY:=VK_TAB ;
 VK_ESCAPE : begin
              Key:=0 ;
              Close ;
              exit;
             end ;
{AA}    65 : if ((Shift=[ssAlt]) and (BVentil.Enabled)) then BEGIN Key:=0 ; BVentilClick(Nil) ; END ;
{AE}    69 : if Shift=[ssAlt] then BEGIN Key:=0 ; BEcheClick(Nil) ; END ;
{AG}    71 : if Shift=[ssAlt] then BEGIN Key:=0 ; BGuideClick(Nil) ; END ;
{AP}    80 : if Shift=[ssCtrl] then BEGIN Key:=0 ; BImprimerClick(Nil) ; END ;
    END ;
end;

procedure TFGuide.BVentilClick(Sender: TObject);
Var T : TVentGuide ;
begin
if OkVentil(FListe.Cells[GCGen,FListe.Row],'')then
   BEGIN
   T:=TVentGuide(FListe.Objects[0,FListe.Row]) ;
   if T=NIL then BEGIN T:=TVentGuide.Create ; FListe.Objects[0,FListe.Row]:=T ; END ;
   ParamGuideAna(T,FListe.Cells[GCGen,FListe.Row],TypeGuide);
   IsModified:=true;
   END else
   BEGIN
   SetFoc(GCGen,FListe.Row) ;
   Messages.Execute(16,'','') ;
   END;
FListe.SetFocus ;
end;

procedure TFGuide.FDeviseChange(Sender: TObject);
begin
IsModified:=true;
DEV.Code:=FDevise.Value ; GetInfosDevise(DEV) ;
end;

procedure TFGuide.BEcheClick(Sender: TObject);
Var St  : string;
begin
if JalFerme then Exit ;
HMrg.Libelle:=FListe.Cells[GCMrg,FListe.Row] ;
St:=FicheRegle_AGL(HMrg.Value,true,Mode);
HMrg.Reload ; HMrg.Value:=St;
IsModified:=true;
FListe.SetFocus ;
end;

procedure TFGuide.BDeleteClick(Sender: TObject);
var ValSt,OldSt : string;
   // ItIndex : integer;
begin
OldSt:=FNomGuide.Text ;
if FNomGuide.Value='' then BEGIN FNomGuide.Text:=OldSt ; Exit ; END ;
if Messages.Execute(18,'Guide : '+FNomGuide.Text+'. ','')<>mrYes then exit;
ValSt:=FNomGuide.Value; //ItIndex:=FNomGuide.ItemIndex ;
ExecuteSQL('DELETE FROM GUIDE WHERE GU_TYPE="'+TypeGuide+'" AND GU_GUIDE="'+ValSt+'"') ;
ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="'+TypeGuide+'" AND EG_GUIDE="'+ValSt+'"') ;
ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="'+TypeGuide+'" AND AG_GUIDE="'+ValSt+'"') ;
FListe.VidePile(TRUE);
ClearScreen; AvertirTable('ttGuideEcr') ;
FNomGuide.Reload ; FNomGuide.Value:='' ;
FDateCreation.Caption:='' ; FDateModification.Caption:='' ; HMrg.Visible:=False ;
Application.ProcessMessages ;
IsModified:=false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 24/05/2002
Modifié le ... :   /  /    
Description .. : Ajout du champs indiquant un guide spécifique à la saisie 
Suite ........ : de trésorerie
Mots clefs ... : 
*****************************************************************}
procedure TFGuide.LoadGuide;
Var Q,QA             : TQuery ;
    St,St1,St2,ValSt,StTva : String ;
    ix,NumLigne  : Integer ;
    T                : TVentGuide ;
    lStJal           : String ;
BEGIN
FListe.VidePile(TRUE);
if FNomGuide.ItemIndex<0 then exit ;
ValSt:=FNomGuide.value;
Q:=OpenSQL('SELECT * FROM GUIDE WHERE GU_TYPE="'+TypeGuide+'" AND GU_GUIDE="'+ValSt+'"',TRUE) ;
lStJal := Q.FindField('GU_JOURNAL').AsString ;
FJournal.Value := lStJal ;
FNaturePiece.Value:=Q.FindField('GU_NATUREPIECE').AsString ;
FEtablissement.Value:=Q.FindField('GU_ETABLISSEMENT').AsString ;
FTypeContrePartie.Value:=Q.FindField('GU_TYPECTREPARTIE').AsString ;
FDevise.Value:=Q.FindField('GU_DEVISE').AsString ;
FDateCreation.Caption:=Q.FindField('GU_DATECREATION').AsString ;
FDateModification.Caption:=Q.FindField('GU_DATEMODIF').AsString ;
//LG* 24/05/2002
GuideTreso.Checked:=StringToCheck(Q.FindField('GU_TRESORERIE').AsString) ;
{ GP 27/02/98 : }
If Trim(FDateCreation.Caption)='' Then FDateCreation.Caption:=DateTimeToStr(Date) ;
If Trim(FDateModification.Caption)='' Then FDateModification.Caption:=DateTimeToStr(Date) ;
{ Fin GP 27/02/98 }
Ferme(Q) ;
Q:=OpenSQL('SELECT * FROM ECRGUI WHERE EG_TYPE="'+TypeGuide+'" AND EG_GUIDE="'+ValSt+'" ORDER BY EG_GUIDE, EG_NUMLIGNE',TRUE) ;
While Not Q.EOF do
   BEGIN
   St:=Q.FindField('EG_ARRET').AsString+'                       ' ;
   NumLigne:=Q.FindField('EG_NUMLIGNE').AsInteger ;

   FListe.Cells[0,FListe.RowCount-1]:=IntToStr(NumLigne);

   FListe.Cells[GCGen,FListe.RowCount-1]:=Q.FindField('EG_GENERAL').AsString ;
   FListe.Cells[GCXGen,FListe.RowCount-1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCAux,FListe.RowCount-1]:=Q.FindField('EG_AUXILIAIRE').AsString ;
   FListe.Cells[GCXAux,FListe.RowCount-1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCRef,FListe.RowCount-1]:=Q.FindField('EG_REFINTERNE').AsString ;
   FListe.Cells[GCXRef,FListe.RowCount-1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCLib,FListe.RowCount-1]:=Q.FindField('EG_LIBELLE').AsString ;
   FListe.Cells[GCXLib,FListe.RowCount-1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCDeb,FListe.RowCount-1]:=Q.FindField('EG_DEBITDEV').AsString ;
   FListe.Cells[GCXDeb,FListe.RowCount-1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCCre,FListe.RowCount-1]:=Q.FindField('EG_CREDITDEV').AsString ;
   FListe.Cells[GCXCre,FListe.RowCount-1]:=St[1] ; Delete(St,1,1) ;

   HMrg.Value:=Q.FindField('EG_MODEREGLE').AsString ;
   FListe.Cells[GCMrg,FListe.RowCount-1]:=HMrg.Text ;
   FListe.Cells[GCXMrg,FListe.RowCount-1]:=St[1] ; Delete(St,1,1) ;

   if VH^.OuiTvaEnc then
      BEGIN
      FListe.Cells[GCXEnc,FListe.RowCount-1]:=Q.FindField('EG_TVAENCAIS').AsString ;
      StTva:=Q.FindField('EG_TVA').AsString ; FTva.Value:=StTva ;
      FListe.Cells[GCTva,FListe.RowCount-1]:=FTva.Text ;
      END ;

   Q.Next ;

   T:=TVentGuide(FListe.Objects[0,FListe.RowCount-1]) ;
   if T<>NIL then T.Free ;
   T:=TVentGuide.Create ; FListe.Objects[0,FListe.RowCount-1]:=T ;
   FListe.RowCount:=FListe.RowCount+1 ;
   QA:=OpenSQL('SELECT * FROM ANAGUI WHERE AG_TYPE="'+TypeGuide+'" AND AG_GUIDE="'+ValSt+'" AND AG_NUMLIGNE='+IntToStr(NumLigne)+' ORDER BY AG_GUIDE, AG_NUMLIGNE, AG_NUMVENTIL',TRUE) ;
   While Not QA.EOF do
      BEGIN
      St1:='' ; St2:=QA.FindField('AG_ARRET').AsString+'              ' ;

      St1:=St1+Format_String(QA.FindField('AG_SECTION').AsString,35)+St2[1] ; Delete(St2,1,1) ;
      St1:=St1+Format_String(QA.FindField('AG_POURCENTAGE').AsString,100)+St2[1] ; Delete(St2,1,1) ;
      St1:=St1+Format_String(QA.FindField('AG_POURCENTQTE1').AsString,100)+St2[1] ; Delete(St2,1,1) ;
      St1:=St1+Format_String(QA.FindField('AG_POURCENTQTE2').AsString,100)+St2[1] ; Delete(St2,1,1) ;
      ix:=StrToInt(Copy(QA.FindField('AG_AXE').AsString,2,1)) ;
      T.Ventil[ix].Add(St1) ;
      QA.Next;
      END ;
   Ferme(QA) ;
   END ;
Ferme(Q) ;
OldGuide:=FNomGuide.Value ; IsModified:=false ;

  // MAJ des accès fiche pour si journal fermé / ouvert
  if not Assigned(FSAJAL) or (FSAJAL.journal <> lStJal) then ; // Chargement du journal // FQ13580 SBO 04/10/2005
    begin
    if Assigned(FSAJAL) then
      FSAJAL.Free ;
    FSAJAL := TSAJAL.Create(lStJal,False) ;
    end ;

 self.Caption := TraduireMemoire('Guides de saisie') + ' : ' + FNomGuide.Text ;
 UpdateCaption(self) ;

 ModeConsult( JalFerme ) ;

END;

procedure TFGuide.BInsertClick(Sender: TObject);
var result : integer;
begin
  if IsModified then
    begin
    Result := Messages.Execute(19,'','');
    if Result=mrYes then
      begin
      if Not ClickValide then Exit ;
      end
    else
      if Result=mrCancel then Exit ;
    end ;
  FListe.VidePile(TRUE);
  ClearScreen;
  FListe.Row  := 1 ;
  FEtablissement.Value := VH^.EtablisDefaut ;
  PositionneEtabUser(FEtablissement) ;
  IsModified  := false;
//  FNomGuide.Enabled := True ;
  if FNomGuide.CanFocus
    then FNomGuide.SetFocus ;
end;

procedure TFGuide.VerifGen(Cpte : String ; R : Integer) ;
var iaxe,isec    : integer;
    ExistV : boolean;
    T            : TVentGuide;
begin
if OldCpte='' then exit ; if OldCpte=Cpte then exit ;
ExistV:=false; 
T:=TVentGuide(FListe.Objects[0,R]) ;
if T=NIL then exit ;
For iaxe:=1 to MaxAxe do
    BEGIN
    For isec:=0 to T.Ventil[iaxe].Count-1 do
      if Trim(T.Ventil[iaxe].Strings[isec])<>'' then  BEGIN ExistV:=true; break; END;
    if ExistV then break;
    END;
if ExistV then
   if Messages.Execute(33,'','')<>mrYes then
      For iaxe:=1 to MaxAxe do
          For isec:=0 to T.Ventil[iaxe].Count-1 do T.Ventil[iaxe].Clear;
end;

procedure TFGuide.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var result : integer;
begin
if IsModified then
   BEGIN
   Result:=Messages.Execute(19,'','');
   if Result=mrYes then BEGIN if Not ClickValide then CanClose:=False ; END else
   if Result=mrCancel then CanClose:=False;
   END;
end;

procedure TFGuide.ClearScreen;
var i,p   : integer;
    Panel : TPanel;
BEGIN
For p:=1 to 2 do
   BEGIN
   if p=1 then Panel:=PGuide else Panel:=PEntete;
   For i:=0 to Panel.ControlCount-1 do
      BEGIN
      if Panel.Controls[i] is TEdit then TEdit(Panel.Controls[i]).Text:='';
      if Panel.Controls[i] is THValComboBox then
         BEGIN
         THValComboBox(Panel.Controls[i]).Text:=''; THValComboBox(Panel.Controls[i]).ItemIndex:=-1;
         END;
      if Panel.Controls[i] is TCheckBox then TCheckBox(Panel.Controls[i]).Checked:=false;
      END;
   END;
END;

procedure TFGuide.BGuideClick(Sender: TObject);
{$IFDEF COMPTA}
Var M : RMVT ;
    DateCpt : TDateTime ;
{$ENDIF}
begin
{$IFDEF COMPTA}
if Not BGuide.Enabled then Exit;
if Mode=taConsult then Exit ;
if JalFerme then Exit ;
if FJournal.Value='' then BEGIN Messages.Execute(36,'','') ; Exit ; END ;
GeneResult:=True ;
Generesult:=ClickValide ; if Not GeneResult then Exit ;
Application.ProcessMessages ;
FillChar(M,Sizeof(M),#0) ;
M.TypeGuide:=TypeGuide ; M.LeGuide:=FNomGuide.Value ; M.FromGuide:=True ;
M.Simul:='S' ; M.Etabl:=FEtablissement.Value ;
M.Jal:=FJournal.Value ; M.DateC:=V_PGI.DateEntree ;
M.DateTaux:=M.DateC ; DateCpt:=M.DateC ;
M.Exo:=QuelExoDT(M.DateC) ;
M.CodeD:=FDevise.Value ; M.TauxD:=GetTaux(M.CodeD,DateCpt,M.DateC) ;
M.Nature:=FNaturePiece.Value ;
LanceSaisie(Nil,taCreat,M) ;
FListe.SetFocus ;
{$ENDIF}
end;

procedure TFGuide.AssisteGuide ;
Var St : String ;
begin
if JalFerme then Exit ;
if (FListe.Col Mod 2)<>0 then Exit ;
if ((FListe.Col<=GCAux) or (FListe.Col>=GCTva)) then Exit ;
St:=ChoixChampZone(FListe.Row,'GUI') ; if St='' then Exit ;
FListe.Cells[FListe.Col,FListe.Row]:=FListe.Cells[FListe.Col,FListe.Row]+St ;
FListe.SetFocus ;
end;

procedure TFGuide.BAssistantClick(Sender: TObject);
begin
AssisteGuide ;
FListe.SetFocus ;
end;

procedure TFGuide.POPSPopup(Sender: TObject);
begin
InitPopup(Self) ;
end;

procedure TFGuide.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  PrintDBGrid (Caption, 'PRT_GUIDE', 'AND GU_TYPE="' + TypeGuide + '"' );
{$ELSE}
  XX_WHERE.Hint := ' Where GU_TYPE="' + TypeGuide + '" ' ;
  PrintDBGrid(Nil,XX_WHERE,Caption,'PRT_GUIDE') ;
{$ENDIF}

end;

procedure TFGuide.FNaturePieceChange(Sender: TObject);
begin
IsModified:=true;
end;

procedure TFGuide.FTypeContrePartieChange(Sender: TObject);
begin
IsModified:=true;
end;

procedure TFGuide.FEtablissementChange(Sender: TObject);
begin
IsModified:=true;
end;


procedure TFGuide.FormClose(Sender: TObject; var Action: TCloseAction);
begin
PurgePopup(POPS) ;
if TypeGuide='ABO' then AvertirTable('ttGuideAbo') else
  if ((TypeGuide='ENC') or (TypeGuide='DEC')) then AvertirTable('ttGuideEncDec') ;
AvertirTable('ttGuideEcr') ;
RegSaveToolbarPos(Self,'Guide') ;
if Parent is THPanel then
   BEGIN
   _Bloqueur('nrSaisieModif',False) ;
   Action:=caFree ;
   END ;
end;

Function TFGuide.ChercheNat( Cpte,Table : string ) : String;
var Q                 : TQuery;
    SQL,FNature,FCpte : String;
BEGIN
Result:='' ;
if Table='GENERAUX' then
   BEGIN
   FNature:='G_NATUREGENE'; FCpte:='G_GENERAL'
   END else if Table='TIERS' then
   BEGIN
   FNature:='T_NATUREAUXI'; FCpte:='T_AUXILIAIRE'
   END ;
SQL:='Select '+FNature+' From '+Table+' Where '+FCpte+'="'+Cpte+'"';
Q:=OpenSQL(SQL,true);
if Not Q.EOF then Result:=Q.FindField(FNature).AsString ;
Ferme(Q);
END;

procedure TFGuide.FNomGuideClick(Sender: TObject);
var Res : integer;
begin
if FNomGuide.ItemIndex>=0 then
   BEGIN
   if IsModified and (OldGuide<>'') and (FNomGuide.value<>OldGuide) then
      BEGIN
      Messages.Execute(48,'','') ; FNomGuide.value:=OldGuide ; Exit ;
      (**)
      Res:=Messages.Execute(19,'','');
      if Res=mrYes then
         BEGIN
         if Not ClickValide then BEGIN FNomGuide.value:=OldGuide ; Exit ; END ;
         END else if Res=mrCancel then FNomGuide.value:=OldGuide ;
      END;
   LoadGuide;
   if Mode=taConsult then
      BEGIN
      FicheReadOnly(Self) ;
//      FNomGuide.Enabled:=True ;
      GeleBoutonPourConsult ;
      END ;
   END else IsModified:=true ;

{$IFDEF EAGLCLIENT}
  // On passe par le mul pour sélectionner le guide en eAGL
//   FNomGuide.Enabled := False ;
{$ELSE}
{$ENDIF}

end;

procedure TFGuide.FNomGuideChange(Sender: TObject);
begin
//if FNomGuide.ItemIndex<0 then IsModified:=true ;
end;

Procedure TFGuide.GeleBoutonPourConsult ( vBoOn : Boolean ) ;
BEGIN
  BInsLigne.Enabled  := not vBoOn ;
  BDelLigne.Enabled  := not vBoOn ;
  BEche.Enabled      := not vBoOn ;
  BVentil.Enabled    := not vBoOn ;
  BAssistant.Enabled := not vBoOn ;
  BGuide.Enabled     := not vBoOn ;
  BInsert.Enabled    := not vBoOn ;
  BDelete.Enabled    := not vBoOn ;
END ;

procedure TFGuide.FListeDblClick(Sender: TObject);
begin
if JalFerme then Exit ;
Case FListe.Col of
   GCGen : BEGIN
           ChercheGen ;
           END ;
   GCAux : BEGIN
           ChercheAux ;
           END ;
   GCXGen,GCXAux,GCXRef,GCXLib,GCXDeb,GCXCre,GCXMrg,GCXEnc :
           BEGIN
           if UpperCase(FListe.Cells[FListe.Col,FListe.Row])='X' then
              BEGIN
              FListe.Cells[FListe.Col,FListe.Row]:='-' ;
              END else
              BEGIN
              FListe.Cells[FListe.Col,FListe.Row]:='X' ;
              END ;
           END ;
   GCTva : TvaSuiv(True) ;
   Else AssisteGuide ;
   END ;
end;

procedure TFGuide.FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
HG.Text:=FListe.Cells[GCGen,FListe.Row] ;
if HG.Text<>'' then BVentil.Enabled:=(HG.ExisteH>0) else BVentil.Enabled:=false;
//BNouveau.Enabled:=(IsModified=false);
//BDelete.Enabled:=(IsModified=false);
if Mode=taConsult then BVentil.Enabled:=False ;
end;

procedure TFGuide.FListeCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
if FListe.Col=GCGEn then OldCpte:=FListe.Cells[GCGen,FListe.Row] else
 if ((Not VH^.OuiTvaEnc) or (TypeGuide='ENC') or (TypeGuide='DEC')) then
    BEGIN
    if ((FListe.Col=GCXEnc) or (FListe.Col=GCTva)) then Cancel:=True ;
    END ;
end;

procedure TFGuide.FormateMontant ( ACol,ARow : Longint ) ;
Var X : Double ;
    St : String ;
BEGIN
St:=FListe.Cells[ACol,ARow] ; if St='' then Exit ;
X:=Valeur(St) ; if X=0 then Exit ;
if Not IsNumeric(St) then Exit ; 
St:=StrfMontant(X,15,DEV.Decimale,'',True) ;
FListe.Cells[ACol,ARow]:=St ;
END ;

procedure TFGuide.FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
Var St : String ;
begin
St:=FListe.Cells[ACol,ARow] ;
if Length(St)>100 then
   BEGIN
   Messages.Execute(47,'','') ; St:=Copy(St,1,100) ;
   FListe.Cells[ACol,ARow]:=St ; Cancel:=True ; Exit ;
   END ;
if ACol=GCGen then
  begin
  VerifGen(FListe.Cells[GCGen,ARow],ARow) ;
  end
else if ACol=GCTva then VerifTva(ACol,ARow) else
   BEGIN
   if Acol in [GCXGen,GCXAux,GCXRef,GCXLib,GCXDeb,GCXCre,GCXMrg] then
      BEGIN
      if uppercase(FListe.Cells[Acol,ARow])<>'X' then FListe.Cells[Acol,ARow]:='-' ;
      END else
      BEGIN
      if ACol=GCXEnc then
         if ((FListe.Cells[Acol,ARow]<>'X') and (FListe.Cells[Acol,ARow]<>'-')) then FListe.Cells[Acol,ARow]:='' ;
      END ;
   END ;
if ACol in [GCDeb,GCCre] then FormateMontant(ACol,ARow) ;
end;

procedure TFGuide.FListeSetEditText(Sender: TObject; ACol, ARow: Longint;const Value: string);
begin
if JalFerme then Exit ;
IsModified:=true ;
//BValider.Enabled:=(IsModified=true);
if ((FListe.Row=FListe.RowCount-1) and (Not G_LigneVide(FListe,FListe.Row))) then
   BEGIN
   if ((TypeGuide='ENC') or (TypeGuide='DEC')) and (FListe.RowCount=3) then else FListe.RowCount:=FListe.RowCount+1 ;
   END;
if ((FListe.Row=FListe.RowCount-3) and (G_LigneVide(FListe,FListe.Rowcount-1))
                                   and (G_LigneVide(FListe,FListe.Rowcount-2))) then
   BEGIN
   if FListe.Objects[0,FListe.Rowcount-1]<>NIL then TVentGuide(FListe.Objects[0,FListe.Rowcount-1]).Free ;
   FListe.RowCount:=FListe.RowCount-1 ;
   END;
G_Renum(FListe) ;
end;

procedure TFGuide.Timer1Timer(Sender: TObject);
begin
//BValider.Enabled:=IsModified ;
end;

procedure TFGuide.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFGuide.FormCreate(Sender: TObject);
begin
RegLoadToolbarPos(Self,'Guide') ;
end;

procedure TFGuide.FListeKeyPress(Sender: TObject; var Key: Char);
begin
if FListe.Col<>GCTva then Exit ;
if Key=' ' then BEGIN Key:=#0 ; TvaSuiv(True) ; END else
   if Key in ['a'..'z','A'..'Z','0'..'9',',','.','/','+','-'] then Key:=#0 ;
end;


procedure TFGuide.FListeExit(Sender: TObject);
Var Cancel : boolean ;
    ACol,ARow : Longint ;
begin
Cancel:=False ; ACol:=FListe.Col ; ARow:=FListe.Row ;
FListeCellExit(Nil,ACol,ARow,Cancel) ;
end;

procedure TFGuide.GuideTresoClick(Sender: TObject);
begin
isModified:=True ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 31/08/2004
Modifié le ... :   /  /
Description .. : - LG - 31/08/2004 - FB 14144 -  
Mots clefs ... : 
*****************************************************************}
procedure TFGuide.SetModified(Value: boolean);
begin
 if JalFerme or ( Mode = taConsult ) then Value := False ;
 FIsModified :=  Value ;
 FNomGuide.Enabled:= not FIsModified ;
end;

procedure TFGuide.ModeConsult( vBoOn : boolean );
begin

  FicheReadOnly( Self , vBoOn ) ;
  FNomGuide.Enabled := True ;

  if not vBoOn then
    begin
    FListe.Options := FListe.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
    FListe.MontreEdit ;
    end ;

  if Mode = taConsult then
    begin
    FNomGuide.Style:=csDropDownList ;
    FNomGuide.Value:=QuelGuide ;
    end
  else if JalFerme then
    begin
    BValider.Enabled    := not vBoOn ;
    end ;

  GeleBoutonPourConsult ( vBoOn ) ;

end;

procedure TFGuide.ChercheGen;
var CGen : TGGeneral ;
begin

  HG.Text := uppercase(FListe.Cells[GCGen,FListe.Row]) ;

  if GChercheCompte( HG, Nil ) then
    begin
    CGen    := TGGeneral.Create( HG.Text ) ;

    // Compte interdit ?
    if EstInterdit( FSAJAL.COMPTEINTERDIT, HG.Text, 0) > 0 then
      begin
      FListe.Cells[GCGen,FListe.Row] := '';
      PgiBox('Ce compte général est interdit sur ce journal.', Caption);
      Exit;
      end ;

    // Compte Fermé ?
    if CGen.Ferme then
      begin
      FListe.Cells[GCGen,FListe.Row] := '';
      PgiBox('Ce compte est fermé. Vous ne pouvez plus l''utiliser dans un guide.', Caption);
      Exit;
      end ;

    // compte OK
    FListe.Cells[GCGen,FListe.Row] := HG.Text ;

    end ;

end;

procedure TFGuide.ChercheAux;
var CAux : TGTiers ;
begin

  HA.Text := uppercase(FListe.Cells[GCAux,FListe.Row]) ;

  if GChercheCompte(HA,Nil) then FListe.Cells[GCAux,FListe.Row]:=HA.Text ;
    begin
    CAux    := TGTiers.Create( HA.Text ) ;

    // Compte Fermé ?
    if CAux.Ferme then
      begin
      FListe.Cells[GCAux,FListe.Row] := '';
      PgiBox('Ce compte est fermé. Vous ne pouvez plus l''utiliser dans un guide.', Caption);
      Exit;
      end ;

    // compte OK
    FListe.Cells[GCAux,FListe.Row] := HA.Text ;

    end ;

end;

function TFGuide.JalFerme: Boolean;
begin
  result := False ;
  if assigned(FSAJAL) then
    result := FSAJAL.OkFerme ;
end;

{function TFGuide.OnSauve: boolean;
var
  Rep : Integer;
begin
  result := FALSE;
  if IsModified then nextPrevcontrol(Self);
  if IsModified then
   begin
    Rep := Messages.Execute(19,'','') ;
    case rep of
      mrYes: if not Bouge(nbPost) then exit;
      mrNo:
        begin
          if (not FClosing) then
            if not Bouge(nbCancel) then exit;
        end;
    mrCancel:
      begin
        Modifier := True;
        Exit;
      end;
  end;
  result := TRUE;
  Modifier := False;
end;  }



function TFGuide.Bouge(Button: TNavigateBtn): boolean;
var
 lInIndex : integer ;
 Rep : integer ;
begin
  result   := false ;
  lInIndex := FNomGuide.ItemIndex ;
  if IsModified then
   begin
    //nextPrevcontrol(Self);
    Rep := Messages.Execute(19,'','');
    if Rep<>mrYes then exit ;
    if not ClickValide then exit ;
   end ;

  case Button of
    nblast : FNomGuide.ItemIndex := FNomGuide.Items.Count - 1 ;
    nbprior : if lInIndex > 0 then FNomGuide.ItemIndex := lInIndex - 1 ;
    nbnext : if lInIndex <= FNomGuide.Items.Count then FNomGuide.ItemIndex := lInIndex + 1 ;
    nbfirst : FNomGuide.ItemIndex := 0 ;
{    nbPost : if not EnregOK then
             begin
               result := FRienAFaire;
               Exit;
             end;  }
  end;

 FNomGuideClick(nil) ;

end;


procedure TFGuide.BNextClick(Sender: TObject);
begin
  Bouge(nbNext);
end;

procedure TFGuide.BPrevClick(Sender: TObject);
begin
 Bouge(nbPrior);
end;

procedure TFGuide.BFirstClick(Sender: TObject);
begin
 Bouge(nbFirst);
end;

procedure TFGuide.BLastClick(Sender: TObject);
begin
 Bouge(nbLast);
end;

end.
