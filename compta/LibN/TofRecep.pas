{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 10/02/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit TofRecep ;

interface

uses Classes,SysUtils,Controls,StdCtrls,ExtCtrls,
     Filectrl, // DirectoryExists
     Hctrls,   // Contrôles Halley
     EtbUser,  // z_GetDestinataires, z_Teletransmission
     hmsgbox,
     Utof,
     ParamSoc,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     PrintDBG,
{$ELSE}
     MaineAGL,
     UtileAGL,
{$ENDIF}
     Ent1,
     HTB97,
     Vierge;

procedure CPLanceFiche_RecepETEBAC3;

type
  TOF_RlvReception = class(TOF)
    G: THGrid;
    cDest : THValComboBox ;
    RRecep : TRadioGroup ;
    BHelp: TToolbarButton97;
  private
    function RecupCarteAppel : string;
    procedure BImprimerClick(Sender: TObject);
    procedure BHelpClick(Sender: TObject);
  public
    procedure Onload ; override ;
    procedure OnUpdate; override ;
  end;

const SR_DEST       =   0 ;
      SR_CART       =   1 ;
      SR_RETO       =   2 ;
      TMsg: array[1..5] of string 	= (
      {01}        '1;Reception de fichiers ETEBAC3;Voulez-vous lancer la reception de fichiers;Q;YN;Y;Y'
      {02}       ,'1;Reception de fichiers ETEBAC3;Le nom du repertoire ETEBAC3 n''est pas valide dans les paramètres société;W;O;O;O'
      {03}       ,'1;Reception de fichiers ETEBAC3;Carte d''appel invalide;W;O;O;O'
      {04}       ,'1;Reception de fichiers ETEBAC3;Le module Liaison Bancaire n''est pas sérialisé ;W;O;O;O;'
      {05}       ,''
      );

implementation

procedure CPLanceFiche_RecepETEBAC3;
begin
  AGLLanceFiche('CP', 'RLVRECEPTION', '', '', '') ;
end;

procedure TOF_RlvReception.OnLoad ;
var BImprimer:TButton ;
begin
TFVierge(Ecran).HelpContext:=7772000;
RRecep:=TRadioGroup(GetControl('REMET')) ;
BImprimer:=TButton(GetControl('BIMPRIME')) ;
BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
if (BImprimer <> nil ) and (not Assigned(BImprimer.OnClick)) then BImprimer.OnClick:=BImprimerClick;
if (BHelp <> nil) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick ;
G:=THGrid(GetControl('GFILES'));
cDest:=THValCombobox(GetControl('CDEST')) ;
cDest.items.add('<<Tous>>');
if cDest<>nil then z_GetDestinataires(cDest.Values,cDest.Items) ;
cDest.ItemIndex:=0 ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 09/04/2002
Modifié le ... :   /  /    
Description .. : - LG* - 09/04/2002 - suppression de la fonction 
Suite ........ : RecupMessageAlmacon 
Mots clefs ... : 
*****************************************************************}
procedure TOF_RlvReception.OnUpdate;
var NomFichier,Chemin,Carte,PDest,Msg : string ;Cpt,i,RetTele,CGrid : integer ;Okok:Boolean ;
begin
if not VH^.OkModEtebac then
  begin
  HShowMessage(TMsg[4],'','') ;
  Exit;
  end ;
Cpt:=cDest.itemindex ;  if Cpt=0 then Cpt:=1;
CGrid:=G.RowCount-1 ;
for i:=Cpt to cDest.items.count-1 do
  begin
  if ((i=cDest.ItemIndex) or (cDest.ItemIndex=0)) and (Cpt<>0) then
    begin
    Okok:=TRUE ; Nomfichier:='FUGIER.ETB' ;RetTele:=-1;
    PDest:=cDest.Values[i-1] ; //Items.Strings[i] ;
    Carte:=RecupCarteAppel ;
    Chemin:=GetParamSocSecur('SO_REPETEBAC3','');
    if Carte='' then begin Msg:='Erreur : Carte d''appel'; Okok:=FALSE; end ;
    if not DirectoryExists(Chemin) then begin Msg:='Erreur : Repertoire ETEBAC3'; Okok:=FALSE; end ;
    if Okok then
      begin
      RetTele:=z_Teletransmission(PChar(NomFichier),PChar(Chemin),PChar(Carte),PChar(PDest),1) ;
      //Msg:=RecupMessageAlmacom(RetTele) ;
      //LG* 
      Msg:='Erreur n°'+IntToStr(RetTele) ;
      end;
    G.Cells[SR_DEST,CGrid]:=PDest;
    G.Cells[SR_CART,CGrid]:=Carte;
    if RetTele = 0 then G.Cells[SR_RETO,CGrid]:='OK' else G.Cells[SR_RETO,CGrid]:=Msg ;
    CGrid:=CGrid+1 ;   G.RowCount:=G.RowCount+1 ;
    end;
  if cDest.ItemIndex<>0 then break ;
  end;
end;

function TOF_RlvReception.RecupCarteAppel : string;
begin
case RRecep.ItemIndex of
  0 : Result:='R000' ;
  1 : Result:='R001' ;
  2 : Result:='R002' ;
  else Result:='';
  end;
end;

procedure TOF_RlvReception.BImprimerClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  PrintDBGrid(G,Nil,'','');
{$ELSE}
  PrintDBGrid(Ecran.Caption,G.ListeParam,'','');
{$ENDIF}
end;

procedure TOF_RlvReception.BHelpClick(Sender: TObject);
begin CallHelpTopic(Ecran) ; end;

Initialization
registerclasses([TOF_RlvReception]);
end.
