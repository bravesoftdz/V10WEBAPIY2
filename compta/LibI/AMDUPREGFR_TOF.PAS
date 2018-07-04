{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 16/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMDUPREGFR ()
Mots clefs ... : TOF;AMDUPREGFR
*****************************************************************}
Unit AMDUPREGFR_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     fe_main,
     mul,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HTB97,
     HEnt1,
     HMsgBox,
     Vierge,
     AGLInit,
     UTOF,
     ImmoREGFR_TOM ;

Type
  TOF_AMDUPREGFR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure OnbFermeClick(Sender : TObject);
    function  ControleZoneOK : boolean;
    procedure EnregistreImmoREGFR;
  end ;


function AMLanceFiche_DupREGFR (Lequel: string): integer;


Implementation

const
  HM: array[1..4] of string =
{1}  ('Le code de destination doit être renseigné.'
{2}   ,'Le libellé de destination doit être renseigné.'
{3}   ,'Le code immobilisation existe déjà.'
{4}   ,'Voulez-vous enregistrer cette opération ?'
      );


function AMLanceFiche_DupREGFR (Lequel: string): integer;
 begin
   result := StrToInt(AGLLanceFiche('AM', 'AMDUPREGFR', '', Lequel, Lequel));
 end;


procedure TOF_AMDUPREGFR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMDUPREGFR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMDUPREGFR.OnUpdate ;
begin
  Inherited ;
  // Ne pas sortir a priori
  Ecran.ModalResult := mrNone;

  if ControleZoneOK then
     EnregistreImmoREGFR;
end ;

procedure TOF_AMDUPREGFR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AMDUPREGFR.OnArgument (S : String ) ;
var T : Tob;
begin
  Inherited ;
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);

  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;

  T := TOB.Create ('IMMOREGFR', nil, -1);
  T.SelectDB ('"'+ S +'"', nil);
  SetControlText('CodeOrig', T.GetValue ('IR_compteimmo'));
  SetControlText('LibOrig', T.GetValue ('IR_LIBELLE'));
  SetControlText('DateOrig', T.GetValue ('IR_DATEACHAT'));
  SetControlText('VOOrig', T.GetValue ('IR_VO') );
  // Préaffichage code & libellé de la nelle immo
  SetControlText('CODE', AMNouveauCodeImmoREGFR);
  SetControlText('LIB', T.GetValue ('IR_LIBELLE'));
  T.Free;
end ;

procedure TOF_AMDUPREGFR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMDUPREGFR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMDUPREGFR.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_AMDUPREGFR.OnbFermeClick(Sender: TObject);
begin
    // Action Echap ou Fermer par la croix
    case PGIAskCancel(HM[4], Ecran.Caption) of
    mrYes:
        // contrôle validation et sortie
        TToolbarButton97(GetControl('BValider')).Click;
    mrNo:
        // Sortir
        Ecran.ModalResult := mrNo;
    else
        // Rester dans l'écran
        Ecran.ModalResult := mrNone;
    end;
end;


function TOF_AMDUPREGFR.ControleZoneOK : boolean;
begin
 result := False;

 if (GetControlText('CODE') = '' ) or (GetControlText('CODE') = '0000000000') then
 begin
    PGIBox(HM[1], Ecran.Caption);
    SetFocusControl('CODE');
 end else
 if (GetControlText('LIB') = '' ) then
 begin
    PGIBox(HM[2], Ecran.Caption);
    SetFocusControl('LIB');
 end else
 if ExisteSQL('SELECT IR_COMPTEIMMO FROM IMMOREGFR WHERE IR_COMPTEIMMO="'+GetControlText('CODE')+'"') then
 begin
    PGIBox(HM[3], Ecran.Caption);
    SetFocusControl('CODE');
 end else
    result := True;
end;

procedure TOF_AMDUPREGFR.EnregistreImmoREGFR;
var T1, T2 : Tob;
begin
  // T1 = enreg origine
  T1 := TOB.Create('IMMOREGFR', nil, -1) ;
  // T2 = enreg destination
  T2 := TOB.Create('IMMOREGFR', nil, -1) ;

  try
    T1.SelectDB ('"'+ GetControlText('CodeOrig') +'"', nil);
    T2.Dupliquer (T1, False, True, True);
    T2.PutValue ('IR_COMPTEIMMO', GetControlText ('CODE'));
    T2.PutValue ('IR_LIBELLE', GetControlText ('LIB'));
    T2.InsertDB (nil);

  finally
    T1.Free ;
    T2.Free ;
    // Sortie OK
    Ecran.ModalResult := mrYes;
    TFVierge(Ecran).retour := IntToStr(mrYes);
    // Retourner le nouveau code immo à l'appelant
    // LaTob <=> TheTob tob globale pointe depuis l'appelant
    if (LaTob <> nil) then
       LaTob.AddChampSupValeur ('CODE', GetControlText ('CODE'));
  end;
end;

Initialization
  registerclasses ( [ TOF_AMDUPREGFR ] ) ;
end.
