{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFormule ()
Mots clefs ... : TOF;AFFormule
*****************************************************************}
Unit UtofAFFormule ;

Interface

Uses StdCtrls,Controls,Classes,Vierge,
{$IFDEF EAGLCLIENT}
    Maineagl,emul,
{$ELSE}
    db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,mul,
{$ENDIF}
   forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,HTB97,UTOF,Paramsoc,Dicobtp ;

Type
  TOF_AFFormule = Class (TOF)
  private
  ListeMontant :TListBox;
  ListeChamp : TStrings;
  FORMULE :THEdit;
  FORMULETEXTE :THLABEL;
  procedure BAjouterClick(Sender: TObject);
  procedure BADDClick(Sender: TObject);
  procedure BMULClick(Sender: TObject);
  procedure BDIVClick(Sender: TObject);
  procedure BEFFClick(Sender: TObject);
  procedure BSOUSClick(Sender: TObject);
  procedure VisuTexteFormule(Sender: TObject);
  public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Procedure AFLanceFiche_FormuleCutOff(Argument:String);

Implementation

Procedure AFLanceFiche_FormuleCutOff(Argument:String);
begin
  AGLLanceFiche ('AFF','AFFORMULECUTOFF','','',Argument);
end;

procedure TOF_AFFormule.OnArgument (S : String ) ;
begin
Inherited ;
  ListeMontant := TListBox (GetControl('LISTEMONTANT'));
  ListeMontant.OnDblClick:=BAjouterClick;
  FORMULE := THEdit (GetControl('FORMULE'));
  FORMULE.onChange :=  VisuTexteFormule;
  FORMULETEXTE := THLABEL (GetControl('FORMULETEXTE'));
  TToolBarButton97 (GetControl('BAJOUTER')).OnClick := BAjouterClick;
  TToolBarButton97 (GetControl('BADD')).OnClick := BADDClick;
  TToolBarButton97 (GetControl('BMUL')).OnClick := BMULClick;
  TToolBarButton97 (GetControl('BDIV')).OnClick := BDIVClick;
  TToolBarButton97 (GetControl('BEFF')).OnClick := BEFFClick;
  TToolBarButton97 (GetControl('BSOUS')).OnClick := BSOUSClick;
end;

procedure TOF_AFFormule.OnLoad ;
begin
  Inherited ;
  ListeChamp := TStringList.Create;
  ListeChamp.Assign ( THValcomboBox(GetControl('CBLISTECHAMP')).Values);
  ListeMontant.Items.Assign ( THValcomboBox(GetControl('CBLISTECHAMP')).Items);
  ListeMontant.Itemindex := 0;
  ListeMontant.setfocus;
  FORMULE.text := GetParamsoc('SO_AFFORMULCUTOFF');
end ;

procedure TOF_AFFormule.OnUpdate ;
begin
  Inherited ;
  SetParamsoc('SO_AFFORMULCUTOFF',FORMULE.text);
end ;

procedure TOF_AFFormule.OnClose ;
begin
  Inherited ;
  ListeChamp.free;
end ;

procedure TOF_AFFormule.BAjouterClick(Sender: TObject);
var i_ind : integer;
    Stchamp :string;
begin
  for i_ind := 0 to ListeMontant.Items.Count - 1 do
    if ListeMontant.Selected[i_ind] then
    begin
      Stchamp := ' ['+ListeChamp[i_ind]+']';
      if (length(FORMULE.text+Stchamp) > 70) then
        PGIInfoAf('Vous dépassez la longueur maximun', Ecran.Caption)
      else
        FORMULE.text := FORMULE.text+' ['+ListeChamp[i_ind]+']';
    end;
end;

procedure TOF_AFFormule.VisuTexteFormule(Sender: TObject);
var stformule:string;
    i_ind :integer;
begin
 stformule :=FORMULE.text;
 for i_ind := 0 to ListeMontant.Items.Count - 1 do
    stformule := findetreplace (stformule,'['+ListeChamp[i_ind]+']',ListeMontant.Items.Strings[i_ind],true);
 FORMULETEXTE.caption := stformule ;
end;

procedure TOF_AFFormule.BADDClick(Sender: TObject);
begin
  FORMULE.text := FORMULE.text + ' +';
end;

procedure TOF_AFFormule.BDIVClick(Sender: TObject);
begin
  FORMULE.text := FORMULE.text + ' /';
end;

procedure TOF_AFFormule.BMULClick(Sender: TObject);
begin
  FORMULE.text := FORMULE.text + ' *';
end;

procedure TOF_AFFormule.BEFFClick(Sender: TObject);
begin
  FORMULE.text := '';
end;

procedure TOF_AFFormule.BSOUSClick(Sender: TObject);
begin
  FORMULE.text := FORMULE.text + ' -';
end;

Initialization
  registerclasses ( [ TOF_AFFORMULE ] ) ;
end.
