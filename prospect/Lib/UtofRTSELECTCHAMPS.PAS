{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 10/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTSELECTCHAMPS ()
Mots clefs ... : TOF;RTSELECTCHAMPS
*****************************************************************}
Unit UtofRTSELECTCHAMPS ;

Interface

Uses StdCtrls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     fe_main,
{$else}
     maineagl,
{$ENDIF}
     HEnt1,
     HMsgBox, 
     UTOF,
     htb97
     ;

Type
  TOF_RTSELECTCHAMPS = Class (TOF)
    private
      LibellesCondi,LibellesDispo : TStrings;
      ChpsDispo,ChpsCondi: TListBox;
      BValide : Boolean;
      procedure AjouterClick(Sender: TObject);
      procedure EnleverClick(Sender: TObject);
      procedure BasculeChampDispo(Champ : string; Indice : integer);
      procedure BasculeChampCondi(Champ : string; Indice : integer);
    public
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
  end ;

Type
     Par_SelectChamps = class
       ListeChamps,ListeChampsSelect : TStringList;
       ListeLibelles : TStrings;
     public
       constructor create;
     end;

procedure RTSelectChamps(SelectChamps : Par_SelectChamps);
Function RTLanceFiche_RTSelectChamps(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
Implementation

var ChampsDispo,ChampsCondi: TStringList;
    LibellesDispoTransit : TStrings;

constructor  Par_SelectChamps.create ;
begin
inherited create;
end;

procedure RTSelectChamps(SelectChamps : Par_SelectChamps);
begin
ChampsDispo := SelectChamps.ListeChamps;
LibellesDispoTransit := SelectChamps.ListeLibelles;
ChampsCondi := SelectChamps.ListeChampsSelect;
AGLLanceFiche('RT','RTSELECTCHAMPS','','','');
end;

procedure TOF_RTSELECTCHAMPS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTCHAMPS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTCHAMPS.OnUpdate ;
begin
  Inherited ;
  BValide := True;
end ;

procedure TOF_RTSELECTCHAMPS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTCHAMPS.OnArgument (S : String ) ;
var iChamp : integer;
begin
  Inherited ;
  TToolBarButton97(GetControl('BAJOUTER')).OnClick:=AjouterClick;
  TToolBarButton97(GetControl('BENLEVER')).OnClick:=EnleverClick;
  ChpsDispo:=TListBox(GetControl('CHPSDISPO'));
  LibellesDispo:=ChpsDispo.Items;
  for iChamp := 0 to LibellesDispoTransit.count-1 do
    begin
    LibellesDispo.add(LibellesDispoTransit.Strings[iChamp]);
    end;
  ChpsCondi:=TListBox(GetControl('CHPSCONDI'));
  LibellesCondi := ChpsCondi.Items;
  BValide := False;
  while ChampsDispo.Count > 0 do
    BasculeChampDispo(ChampsDispo.Strings[0],-1);
end ;

procedure TOF_RTSELECTCHAMPS.OnClose ;
begin
  Inherited ;
  if BValide = False then ChampsCondi.clear;
end ;

procedure TOF_RTSELECTCHAMPS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTCHAMPS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_RTSELECTCHAMPS.AjouterClick(Sender: TObject);
begin
  if ChpsDispo.Items.Count > 0  then
     begin
     if (ChpsDispo.ItemIndex = -1 ) or ( (ChpsDispo.ItemIndex >= 0) and
        ( ChpsDispo.Selected[ChpsDispo.ItemIndex] = False) ) then
         begin
         PGIBox(TraduireMemoire('Vous devez sélectionner un champ disponible'),'Ajouter un champ');
         exit;
         end;
     BasculeChampDispo(ChampsDispo.Strings[ChpsDispo.ItemIndex],ChpsDispo.ItemIndex);
     end;
end;

procedure TOF_RTSELECTCHAMPS.EnleverClick(Sender: TObject);
begin
  if ChpsCondi.Items.Count > 0  then
     begin
     if (ChpsCondi.ItemIndex = -1  ) or ( (ChpsCondi.ItemIndex >= 0) and
          ( ChpsCondi.Selected[ChpsCondi.ItemIndex] = False) ) then
        begin
        PGIBox(TraduireMemoire('Vous devez sélectionner un champ sélectionné'),'Enlever un champ');
        exit;
        end;
        BasculeChampCondi(ChampsCondi.Strings[ChpsCondi.ItemIndex],ChpsCondi.ItemIndex)
     end;
end;

procedure TOF_RTSELECTCHAMPS.BasculeChampDispo(Champ : string; Indice : integer);
var Posit : integer;
begin
  // function Find(const S: string; var Index: Integer): Boolean; virtual;
  //if not ChampsDispo.find(Champ,Posit) then exit;
  if Indice = -1 then
    begin
    for Posit:=0 to ChampsDispo.Count-1 do
       if ChampsDispo.Strings[Posit]=Champ then break;
    if Posit >= ChampsDispo.Count then exit;
    end
  else Posit := Indice;

  ChampsCondi.add (Champ);
  LibellesCondi.add (LibellesDispo.Strings[Posit]);
  ChampsDispo.Delete(Posit);
  LibellesDispo.Delete(Posit);
end;

procedure TOF_RTSELECTCHAMPS.BasculeChampCondi(Champ : string; Indice : integer);
var Posit : integer;
begin
  // function Find(const S: string; var Index: Integer): Boolean; virtual;
  { bien tenté mais j'ai enlevé le sort ...if not ChampsCondi.find(Champ,Posit) then exit;}
  if Indice = -1 then
    begin
    for Posit:=0 to ChampsCondi.Count-1 do
       if ChampsCondi.Strings[Posit]=Champ then break;
    end
  else Posit := Indice;
  if Posit >= ChampsCondi.Count then exit;
  ChampsDispo.add (Champ);
  LibellesDispo.add (LibellesCondi.Strings[Posit]);
  ChampsCondi.Delete(Posit);
  LibellesCondi.Delete(Posit);
end;

Function RTLanceFiche_RTSelectChamps(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

Initialization
  registerclasses ( [ TOF_RTSELECTCHAMPS ] ) ; 
end.
