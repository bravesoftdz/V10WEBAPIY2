{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /    
Description .. : Unit gestion du multi critère gestion des contrats de travail
Mots clefs ... : PAIE;CONTRAT
*****************************************************************}
{
PT1  : 23/09/02   : V585  PH Désactivation du filtre par défaut
PT2  : 19/12/2007 FC V810 FQ 14996 Concepts accessibilité depuis la fiche salarié
PT3  : 22/01/2008 FC V_81 FQ 14147 Pb en insert quand pas de contrat existant
PT4  : 01/04/2008 FC V_90 Pb rafraichissement mul contrat depuis fiche salarié
}
unit UTofPG_MULCONTRAT;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
       MaineAgl,emul,
{$ENDIF}
      Grids,HCtrls,HEnt1, vierge,EntPaie, HMsgBox,Hqry,UTOF, UTOB, UTOM,
      AGLInit, ParamDat,PgOutils, P5Def;

Type
     TOF_PGMULCONTRAT = Class (TOF)
       private
       LeSalarie : THEdit;
       Action:String;
//       procedure ClickInsert (Sender: TObject);
        procedure DoubleCLickListe(Sender: TObject);   //PT2
        procedure ClickInsert(Sender: Tobject); //PT2
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation

{ Jamais appelée
procedure TOF_PGMULCONTRAT.ClickInsert(Sender: TObject);
var st    : string;
begin
St := LeSalarie.Text+';ACTION=CREATION';
AGLLanceFiche ('PAY','CONTRATTRAVAIL', '','', st);
end;
}

//DEB PT2
procedure TOF_PGMULCONTRAT.ClickInsert(Sender: Tobject);
begin
  AGLLanceFiche ('PAY','CONTRATTRAVAIL', '','', GetControlText('PCI_SALARIE')+';ACTION=CREATION');//PT3
  TFMul(Ecran).BChercheClick(Sender);  //PT4
end;

procedure TOF_PGMULCONTRAT.DoubleCLickListe(Sender: TObject);
begin
  if TFmul(Ecran).Q.RecordCount <> 0 then
//  if GetField('PCI_SALARIE') <> '' then
  begin
    if (not JaiLeDroitTag(200062)) and (Action='ACTION=CONSULTATION') then
      AGLLanceFiche ('PAY','CONTRATTRAVAIL', '',GetField('PCI_SALARIE')+';'+IntToStr(GetField('PCI_ORDRE')), GetField('PCI_SALARIE')+';'+'ACTION=CONSULTATION')
    else
      AGLLanceFiche ('PAY','CONTRATTRAVAIL', '',GetField('PCI_SALARIE')+';'+IntToStr(GetField('PCI_ORDRE')), GetField('PCI_SALARIE')+';'+'ACTION=MODIFICATION');
  end
  else
    if (JaiLeDroitTag(200062)) and (Action='ACTION<>CONSULTATION') then
      AGLLanceFiche ('PAY','CONTRATTRAVAIL', '','', GetField('PCI_SALARIE')+';ACTION=CREATION');
  TFMul(Ecran).BChercheClick(Sender);  //PT4
end;
//FIN PT2

procedure TOF_PGMULCONTRAT.OnArgument(Arguments: String);
var Arg    : String;
{$IFNDEF EAGLCLIENT}
  FListe : THDBGrid;
{$ELSE}
  FListe : THGrid;
{$ENDIF}
  Btn:TToolbarButton97;
begin
inherited ;

  Arg := Arguments;
  Arg := Trim (Arg);
  if Arg = '' then exit;
  TFmul(Ecran).FiltreDisabled := TRUE;
  LeSalarie := THEdit (GetControl ('PCI_SALARIE'));
  if LeSalarie <> NIL then LeSalarie.Text := READTOKENST(Arg); // Recup du code salarie sur lequel on travaille
  //DEB PT2
  Action:='';
  if Arg <> '' then
    Action:=READTOKENST(Arg);
{$IFNDEF EAGLCLIENT}
  FListe := THDBGrid(GetControl('FListe'));
{$ELSE}
  FListe := THGrid(GetControl('FListe'));
{$ENDIF}
  if FListe<>nil then
    Fliste.OnDblClick:=DoubleCLickListe;

  Btn:=TToolbarButton97 (GetControl ('BINSERT'));
  if Btn <> NIL then
    Btn.OnClick := ClickInsert;

  if (Action='ACTION=CONSULTATION') then
    SetControlEnabled('BINSERT',False);
  //FIN PT2
end;


Initialization
registerclasses([TOF_PGMULCONTRAT]);
end.


