{***********UNITE*************************************************
Auteur  ...... : Tynévez Bernadette
Créé le ...... : 17/07/2007
Modifié le ... : 07/08/2007
Description .. : Source TOF de la FICHE : AMARDDOTGLOBALE ()
Suite ........ : 
Mots clefs ... : TOF;AMARDDOTGLOBALE
*****************************************************************}
Unit AMARDDOTGLOBALE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
     HTB97,
     UTOF ;

Type
  TOF_AMARDDOTGLOBALE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    fDotMin : double;
    fDotMax : double;
    fDotTheorique : double;
    procedure OnbFermeClick(Sender: TObject);
    function  ControleZoneOK : boolean;
    procedure EnregistrerDotation;
  end ;

Function AMLanceFiche_ARDDotGlobale(stParam: string='') : integer;


Implementation

const
 HMD: array[0..1] of string =
 {0} ('Voulez-vous enregistrer cette opération ?',
 {1}  'Le montant se situe en dehors de la fourchette autorisée pour les dotations.'
      );


function AMLanceFiche_ARDDotGlobale(stParam: string='') : integer;
begin
 result:= StrToInt(AGLLanceFiche('AM', 'AMARDDOTGLOBALE','' , '', stParam));
end;

procedure TOF_AMARDDOTGLOBALE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMARDDOTGLOBALE.OnDelete ;
begin
  Inherited ;
end ;

// Action Valider l'écran
procedure TOF_AMARDDOTGLOBALE.OnUpdate ;
begin
  Inherited ;
  // Ne pas sortir a priori
  Ecran.ModalResult := mrNone;

  if ControleZoneOK then
     EnregistrerDotation;
end ;

procedure TOF_AMARDDOTGLOBALE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMARDDOTGLOBALE.OnArgument (S : String ) ;
var i:integer;
begin
  Inherited ;
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);
  fDotMin := 0;
  fDotMax := 0;
  fDotTheorique := 0;

  // Les immos sélectionnées sont dans TheTob générale qui s'appelle LaTob ici dans l'appelé
  // Champs de la Tob :
  // - sélection des immos par Ctl+A : ceux de l'OpenSql de RecupereAllSelect de l'appelant
  // - sélection des immos par barre espace : ceux de RecupereSelect de l'appelant
  if (LaTob <> nil) then
     for i:=0 to LaTob.Detail.count-1 do
     begin
        fDotMin := fDotMin + LaTob.Detail[i].GetValue ('I_DOTMIN');
        fDotMax := fDotMax + LaTob.Detail[i].GetValue ('I_DOTMAX');
        fDotTheorique := fDotTheorique + LaTob.Detail[i].GetValue ('I_DOTTHEORIQUE');
     end;

  // Initialisation des champs
  SetControlEnabled('DOTMIN', False);
  SetControlProperty('DOTMIN', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlText('DOTMIN', StrFMontant(fDotMin, 15, V_PGI.OkDecV, '', True));
  SetControlEnabled('DOTMAX', False);
  SetControlProperty('DOTMAX', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlText('DOTMAX', StrFMontant(fDotMax, 15, V_PGI.OkDecV, '', True));

  // Evènements
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
end ;

procedure TOF_AMARDDOTGLOBALE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMARDDOTGLOBALE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMARDDOTGLOBALE.OnCancel () ;
begin
  Inherited ;
end ;

// Action Echap ou Fermer par la croix
procedure TOF_AMARDDOTGLOBALE.OnbFermeClick(Sender: TObject);
begin
    case PGIAskCancel(HMD[0], Ecran.Caption) of
    mrYes:
      begin
      if ControleZoneOK then
           // OK enregistrer et sortir
           EnregistrerDotation
      else Ecran.ModalResult := mrNone; // Rester dans l'écran
      end;
    mrNo:
      // Sortir
      Ecran.ModalResult := mrNo;
    else
      // Rester dans l'écran
      Ecran.ModalResult := mrNone;
    end
end;

function TOF_AMARDDOTGLOBALE.ControleZoneOK : boolean;
begin
 result := True;
 if (Valeur(GetControlText('DotGlobale')) < fDotMin ) or
    (Valeur(GetControlText('DotGlobale')) > fDotMax ) then
    begin
    result := False;
    PGIBox(HMD[1], Ecran.Caption);
    end;
end;

procedure TOF_AMARDDOTGLOBALE.EnregistrerDotation ;
begin
  // Sortie OK
  Ecran.ModalResult := mrYes;
  TFVierge(Ecran).retour := IntToStr (mrYes);

  // Enregistrement en table IMMO à faire
  // => màj Tob inutile, l'appelant en retour relance la requête de multisélection qui rafraichit la grille
  // for i:=0 to LaTob.Detail.count-1 do pour boucler sur chaque immo
  //  LaTob.Detail[i].GetValue ('I_IMMO') donne le code immo à traiter
end;


Initialization
  registerclasses ( [ TOF_AMARDDOTGLOBALE ] ) ;
end.
