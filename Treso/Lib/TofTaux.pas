{***********UNITE*************************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 23/11/2001
Modifié le ... : 05/12/2001
Description .. : Source TOF de la TABLE : TOFTAUX ()
               : JP 03/08/03 : Migration eAGL
Mots clefs ... : TOF;TOFTAUX
*****************************************************************}
Unit TofTaux ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAGL, HCtrls,
  {$ELSE}
  FE_Main, DBGrids, HDB,
  {$ENDIF}
  Classes, HTB97, HEnt1,  AglInit, UTOF, Controls;

Type
  TOF_COTATIONTAUX = Class (TOF)
  private
    BOuvrir     : TToolBarButton97 ;
    BInsert	: TToolBarButton97 ;
    BDelete     : TToolBarButton97 ;
    BFirst      : TToolBarButton97 ;
    BCherche    : TToolBarButton97 ;
    {$IFDEF EAGLCLIENT}
    FListe      : THGrid ;
    {$ELSE}
    FListe      : THDBGrid ;
    {$ENDIF}

    procedure BOuvrirOnClick		(Sender : TObject) ;
    procedure BInsertOnClick		(Sender : TObject) ;
    procedure BDeleteOnClick		(Sender : TObject);

  public
    procedure OnArgument (S : String ) ; override ;
  end ;

procedure TRLanceFiche_CotationTaux(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

Implementation

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_CotationTaux(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_COTATIONTAUX.BOuvrirOnClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  str    : string ;
  strD   : string ;
  phrase : string ;
begin
  strD := VarToStr(GetField('TTA_DATE'));
  str  := VarToStr(GetField('TTA_CODE'));
  phrase := str+ ';' + strD ;
  AglLanceFiche('TR','TRSAISIETAUX','', phrase,'ACTION=MODIFICATION;'+phrase) ;
  if Bcherche <> nil then BCherche.Click ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_COTATIONTAUX.BInsertOnClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche( 'TR','TRSAISIETAUX', '','',ActionToString(taCreat));
  if Bcherche <> nil then BCherche.Click ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_COTATIONTAUX.BDeleteOnClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
begin
  if Bcherche <> nil then BCherche.Click ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_COTATIONTAUX.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  BFirst   := TToolBarButton97(GetControl('BFirst'));
  bOuvrir  := TToolBarButton97(GetControl('BOuvrir'));
  BDelete  := TToolBarButton97(GetControl('BDelete'));
  BInsert  := TToolBarButton97(GetControl('BInsert'));
  BCherche := TToolBarButton97(GetControl('BCherche'));
  {$IFDEF EAGLCLIENT}
  FListe   := THGrid  (GetControl('Fliste'));
  {$ELSE}
  FListe   := THDBGrid(GetControl('Fliste'));
  {$ENDIF}

  if BOuvrir	<> nil then BOuvrir.OnClick   := BOuvrirOnClick;
  if BFirst    	<> nil then BFirst.Visible    := False;
  if FListe	    <> nil then Fliste.OnDblClick := BOuvrirOnClick;
  if BInsert	<> nil then BInsert.OnClick   := BInsertOnClick;
  if BDelete	<> nil then BDelete.OnClick   := BDeleteOnClick;
end;


Initialization
  registerclasses ( [ TOF_COTATIONTAUX ] ) ;
end.

