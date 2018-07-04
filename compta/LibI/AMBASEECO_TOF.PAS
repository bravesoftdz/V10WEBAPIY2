{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 17/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMBASEECO ()
Mots clefs ... : TOF;AMBASEECO
*****************************************************************}
Unit AMBASEECO_TOF ;

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
     uTob,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ,
     HTB97,
     AGLInit, // TheData
     Vierge;

Type
  TOF_AMBASEECO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    public
      fCorrection : double;
      fBaseEco : double;
    private
      fAction : string;
      procedure OnbFermeClick(Sender: TObject);
      procedure OnExitCorrection(Sender: TObject);
      procedure OnChangeCorrection(Sender: TObject);
      //procedure OnExitBaseEco(Sender: TObject);
      function  ControleZoneOK : boolean;
      procedure EnregistrerCorrection;
  end ;

Function AMLanceFiche_BaseEco(stParam: string) : integer;

Implementation

const
 HMD: array[0..1] of string =
 {0} ( 'Voulez-vous modifier la base d''amortissement économique ?',
 {1}   'Montant négatif' );


function AMLanceFiche_BaseEco(stParam: string) : integer;
begin
 result:= StrToInt(AGLLanceFiche('AM', 'AMBASEECO','' , '', stParam));
end;

procedure TOF_AMBASEECO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMBASEECO.OnDelete ;
begin
  Inherited ;
end ;

// Action Valider l'écran
procedure TOF_AMBASEECO.OnUpdate ;
begin
  Inherited ;
  // Ne pas sortir a priori
  Ecran.ModalResult := mrNone;

  if ControleZoneOK then
     EnregistrerCorrection;
end ;

procedure TOF_AMBASEECO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMBASEECO.OnArgument (S : String ) ;
var Theorique : double;
begin
  Inherited ;
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);

  // Récup paramètres en entrée
  fAction := ReadTokenSt(S);
  Theorique := StrToFloat(ReadTokenSt(S));

  // Récup base éco dans TheTob (tob globale) qui s'appelle LaTob ici dans l'appelé
  fBaseeco := 0;
  if (LaTob <> nil) then
     fBaseeco := LaTob.GetValue ('BASEECO') ;
  fCorrection := Arrondi (Theorique-fBaseeco, V_PGI.OkDecV);

  // Initialisation des champs
  SetControlEnabled('MtTheorique', False);
  SetControlProperty('MtTheorique', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('MtCorrection', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('MtBase', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlText('MtTheorique', StrFMontant(Theorique,15,V_PGI.OkDecV,'',True));
  SetControlText('MtCorrection', StrFMontant(fCorrection,15,V_PGI.OkDecV,'',True));
  SetControlText('MtBase', StrFMontant(fBaseeco,15,V_PGI.OkDecV,'',True));
  //if fAction='ACTION=CREATION' then
  SetControlEnabled('MtBase', False);
  //else  SetControlEnabled('MtCorrection', False);

  // Evènements
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
  THEdit(GetControl('MtCorrection')).OnExit    := OnExitCorrection;
  THEdit(GetControl('MtCorrection')).OnChange    := OnChangeCorrection;
  //THEdit(GetControl('MtBase')).OnExit    := OnExitBaseEco;
end ;

procedure TOF_AMBASEECO.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMBASEECO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMBASEECO.OnCancel () ;
begin
  Inherited ;
end ;

// Action Echap ou Fermer par la croix
procedure TOF_AMBASEECO.OnbFermeClick(Sender: TObject);
begin

  if (Valeur(GetControlText('MtCorrection')) <> fCorrection) then
     //(Valeur(GetControlText('MtBase')) <> fBaseEco ) then

    case PGIAskCancel(HMD[0], Ecran.Caption) of
    mrYes:
      begin
      if ControleZoneOK then
           // OK enregistrer et sortir
           EnregistrerCorrection
      else Ecran.ModalResult := mrNone; // Rester dans l'écran
      end;
    mrNo:
      begin
      // Sortir
      Ecran.ModalResult := mrNo;
      end;
    else
      // Rester dans l'écran
      Ecran.ModalResult := mrNone;
    end

  else
  // pas de modifs => sortir
  Ecran.ModalResult := mrNo;

end;

procedure TOF_AMBASEECO.OnChangeCorrection(Sender: TObject);
var theorique, correction : double;
begin
  theorique := Valeur(GetControlText('MtTheorique'));
  correction := Valeur(GetControlText('MtCorrection'));
  SetControlText('MtBase', FloatToStr(Arrondi(theorique-correction, V_PGI.OkDecV)));
end;

procedure TOF_AMBASEECO.OnExitCorrection(Sender: TObject);
var theorique, correction : double;
begin
  theorique := Valeur(GetControlText('MtTheorique'));
  correction := Valeur(GetControlText('MtCorrection'));
  SetControlText('MtBase', FloatToStr(Arrondi(theorique-correction, V_PGI.OkDecV)));
end;

{procedure TOF_AMBASEECO.OnExitBaseEco(Sender: TObject);
var theorique, baseeco : double;
begin
  theorique := Valeur(GetControlText('MtTheorique'));
  baseeco := Valeur(GetControlText('MtBase'));
  SetControlText('MtCorrection', FloatToStr(Arrondi(theorique-baseeco, V_PGI.OkDecV)));
end; }


function TOF_AMBASEECO.ControleZoneOK : boolean;
begin
 // Provoquer le recalcul
 //if THEdit (GetControl('MtCorrection')).Enabled then OnExitCorrection (nil)
 //else OnExitBaseEco (nil);
 OnExitCorrection (nil);
 result := (Valeur(GetControlText('MtCorrection'))>0) or (Valeur(GetControlText('MtCorrection'))=0);
 if not result then
    PGIBox(HMD[1], '')
 else
    begin
    result := (Valeur(GetControlText('MtBase'))>0) or (Valeur(GetControlText('MtBase'))=0);
    if not result then
      PGIBox(HMD[1], '');
    end;
end;

procedure TOF_AMBASEECO.EnregistrerCorrection ;
begin
  Ecran.ModalResult := mrYes;
  TFVierge(Ecran).retour := IntToStr (mrYes);

  // MAJ base éco dans TheTob (tob globale) qui s'appelle LaTob ici dans l'appelé
  if (LaTob <> nil) then
    LaTob.PutValue ('BASEECO', Valeur(GetControlText('MtBase')));
end;

Initialization
  registerclasses ( [ TOF_AMBASEECO ] ) ;
end.
