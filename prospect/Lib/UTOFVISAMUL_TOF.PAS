{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 18/05/2006
Modifié le ... : 22/05/2006
Description .. : Source TOF de la FICHE : UTOFVISAMUL ()
Suite ........ : Liste des ciblages pour confirmation Visa
Mots clefs ... : TOF;UTOFVISAMUL
*****************************************************************}
Unit UTOFVISAMUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     Maineagl,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     utilGc,    //changelibre2
     HMsgBox,
     UTOF,
     HTB97,
     KPMGUtil,
     AssistCiblageVersOperation,
     {$ifdef AFFAIRE}
     UTOFAFTRADUCCHAMPLIBRE,
     {$endif}
     UTOFLISTEPOURVISA_TOF,
     UTOFLISTEVISACON_TOF
      ;

Type
{$ifdef AFFAIRE}
                //PL le 25/05/07 pour gérer les tablettes en correspondance
     TOF_UTOFVISAMUL = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_UTOFVISAMUL = Class (TOF)
{$endif}
  public
    LaListe : THGrid;
    Bselect : TToolBarButton97;
    Boperation : TToolBarButton97;
    procedure BselectOnClick ( Sender : TObject );
    procedure BoperationOnClick ( Sender : TObject );


    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

  Function RTLanceFiche_utofvisamul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_utofvisamul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_UTOFVISAMUL.OnArgument (S : String ) ;
var
//   Qr : TQuery;
   i  : integer;
begin
  Inherited ;

  //SetControlVisible('BOUVRIR', False);

  //Bselect := TToolBarButton97(getcontrol('BSELECT'));
  //Bselect.OnClick := BselectOnClick;

  // PL le 17/04/07 : on inverse le comportement des boutons Select et ouvrir
  SetControlVisible('BSELECT', False);
  Bselect := TToolBarButton97(getcontrol('BOUVRIR'));
  Bselect.OnClick := BselectOnClick;

  LaListe := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick := BselectOnClick;

  Boperation := TToolBarButton97(GetControl('BOPERATION'));
  Boperation.OnClick := BoperationOnClick;

//  Qr := OpenSQL('SELECT MN_LIBELLE FROM MENU WHERE MN_TAG="92265"', True);
//  If not Qr.EOF then
//  begin
//    TFMul(Ecran).Caption := Qr.Fields[0].AsString;
    // PL le 17/04/07 : en dur car le titre est fixe
    TFMul(Ecran).Caption := 'Validation ciblage';
    UpdateCaption(TFmul(Ecran));
//  end;


//  Ferme(Qr);

  // maj des champs libres
   for i := 1 to 5 do
   begin
    ChangeLibre2('TRCB_RCBTABLELIBRE'+intToStr(i), Tform(Ecran));
    ChangeLibre2('TRCB_RCBTEXTELIBRE'+intToStr(i), Tform(Ecran));
    ChangeLibre2('TRCB_RCBDATELIBRE'+intToStr(i), Tform(Ecran));
   end;

end ;

procedure TOF_UTOFVISAMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_UTOFVISAMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_UTOFVISAMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_UTOFVISAMUL.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_UTOFVISAMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_UTOFVISAMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFVISAMUL.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_UTOFVISAMUL.BselectOnClick(Sender: TObject);
begin
if VarIsNull(GetField('RCB_CODECIBLAGE')) then
   exit;

//pas de topage
if Not ExisteSql ('SELECT 1 FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+GetField('RCB_CODECIBLAGE')+'"') then
begin
   PgiBox ('Il n''y a rien à viser !');
   exit;
end;


//si ciblage pour raison sociale
if GetField('RCB_TYPECIBLAGE')='001' then
   RTLanceFiche_utoflistepourvisa('RT', 'RTTOPVISA', '', '', Getfield('RCB_CODECIBLAGE'))
else
   RTLanceFiche_UTOFLISTEVISACON('RT', 'RTTOPVISACONTACT', '', '', Getfield('RCB_CODECIBLAGE'));



//TFmul(ecran).Q.UpdateCriteres;
TFMul(ecran).BChercheClick(Ecran);


end;



procedure TOF_UTOFVISAMUL.BoperationOnCLick(Sender: TObject);
begin
if GetField('RCB_CODECIBLAGE')=null then
   exit;

if not ExisteSQL('SELECT 1 FROM CIBLAGE WHERE RCB_CODECIBLAGE="'+GetField('RCB_CODECIBLAGE')+
                 '" AND RCB_VALIDATION="X"') then
begin
   PGIError('Ce ciblage n''est pas validé');
   exit;
end;



//CiblageVersOperation(Getfield('RCB_CODECIBLAGE'));
Assist_CiblageVersOperation(Getfield('RCB_CODECIBLAGE'));

end;



Initialization
  registerclasses ( [ TOF_UTOFVISAMUL ] ) ;
end.
