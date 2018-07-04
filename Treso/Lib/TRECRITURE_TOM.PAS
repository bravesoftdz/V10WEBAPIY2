{ Unité : Source TOM de la TRECRITURE : écritures validées BO
--------------------------------------------------------------------------------------
    Version  |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91          21/07/03  JP   Création de l'unité
6.xx.xxx.xxx   20/07/04  JP   Gestion des commissions dans la réalisation et l'intégration cf OnAfterUpdateRecord
6.0X.001.001   21/08/04  JP   Correction de la gestion des taux de change et des décimales
6.50.001.001   24/03/05  JP   Dans le cadre de la FQ 10223, je remplace les TrEcritures par des TOB.
                              Mise à jour des clefs si changement des dates
7.05.001.001   13/10/06  JP   Gestion des comptes courants : on peut modifier modifier les dates
8.00.002.002   01/08/07  JP   Nouvelle gestion des soldes.
-------------------------------------------------------------------------------------}

unit TRECRITURE_TOM ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes,
  {$IFDEF EAGLCLIENT}
     eFiche,
  {$ELSE}
     Fiche, HDB, DB, 
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, UTOM, UTob, ExtCtrls;

type
  TOM_TRECRITURE = class (TOM)
    procedure OnArgument   (S : string); override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnLoadRecord             ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnChangeField(F : TField); override; {01/08/07}
  private
    CanModif   : Boolean;
    Taux       : Double;
    NbDec      : Integer; {21/10/04}
    Composant  : string;
    CptCourant : Boolean; {13/10/06}

    procedure DateExit     (Sender : TObject);
    procedure MontantChange(Sender : TObject);
    procedure MontantEnter (Sender : TObject);
    procedure ApresFormShow;
    procedure MajAffichage; {01/08/07}
  end ;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Commun, Constantes, UProcCommission, UProcGen, UProcSolde, Math, Ent1;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  c : string;
  {$IFDEF EAGLCLIENT}
  e : THEdit;
  v : THValComboBox;
  {$ELSE}
  v : THDBValComboBox;
  {$ENDIF}
begin
  inherited;
  Ecran.HelpContext := 150;
  c := ReadTokenSt(s);
  {$IFDEF EAGLCLIENT}
  v := THValComboBox(GetControl('TE_CODEFLUX'));
  {$ELSE}
  v := THDBValComboBox(GetControl('TE_CODEFLUX'));
  {$ENDIF}
  {Pour les rubriques, c'est géré dans le OnLoad}
  if c = na_Simulation then v.DataType := 'TZRUBRIQUE'
                       else v.DataType := 'TRCODEFLUX';

  c := ReadTokenSt(S);
  CptCourant := c = TRANSACICC;
  CanModif   := c <> '';

  if CanModif then begin
    TFFiche(Ecran).OnAfterFormShow := ApresFormShow; {21/10/04 : c'est mieux que le FormActivate}
    TFFiche(Ecran).TypeAction := taModif;
    {$IFDEF EAGLCLIENT}
    e := THEdit(GetControl('TE_DATECOMPTABLE'));
    e.OnExit := DateExit;
    e.OpeType := otDate;
    e.EditMask := '!99/99/0000;1;_';
    {$ENDIF}
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.ApresFormShow;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  c : string;
begin
  SetControlVisible('BINSERT', False);
  SetControlVisible('BDELETE', False);
  SetControlVisible('BDEFAIRE', True);

  for n := 0 to Ecran.ComponentCount - 1 do begin
    c := Uppercase(Ecran.Components[n].Name);
    if (c <> 'TE_DATECOMPTABLE') and (c <> 'TE_MONTANT') and (c <> 'TE_MONTANTDEV') and (c <> 'DEVISE') and
       (c <> 'TTE_DATECOMPTABLE') and (c <> 'TTE_MONTANT') and (c <> 'TTE_MONTANTDEV') and
       (c <> 'PGENERAL') and (c <> 'DOCK971') and (c <> 'PAGES') and (c <> 'PBOUTON') and (c <> 'BDEFAIRE') and
       (c <> 'BIMPRIMER') and (c <> 'BVALIDER') and (c <> 'BFERME') and (c <> 'HELPBTN') then
      SetControlEnabled(c, False);
    {13/10/06 : sur les comptes courants, seules les dates sont modifiables}
    if CptCourant then begin
           if (c = 'TE_DATEVALEUR' ) then SetControlEnabled(c, True)
      else if (c = 'TTE_DATEVALEUR') then SetControlEnabled(c, True)
      else if (c = 'TE_MONTANT'    ) then SetControlEnabled(c, False)
      else if (c = 'TTE_MONTANT'   ) then SetControlEnabled(c, False)
      else if (c = 'TE_MONTANTDEV' ) then SetControlEnabled(c, False)
      else if (c = 'TTE_MONTANTDEV') then SetControlEnabled(c, False);
    end;
  end;

  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('TE_DATECOMPTABLE')).ElipsisButton := True;
  THEdit(GetControl('TE_DATECOMPTABLE')).OnExit   := DateExit;
  THEdit(GetControl('TE_MONTANT'      )).OnChange := MontantChange;
  THEdit(GetControl('TE_MONTANTDEV'   )).OnChange := MontantChange;
  THEdit(GetControl('TE_MONTANT'      )).OnEnter  := MontantEnter;
  THEdit(GetControl('TE_MONTANTDEV'   )).OnEnter  := MontantEnter;
  {$ELSE}
  THDBEdit(GetControl('TE_DATECOMPTABLE')).ElipsisButton := True;
  THDBEdit(GetControl('TE_DATECOMPTABLE')).EditMask := '!99/99/0000;1;_';
  THDBEdit(GetControl('TE_DATECOMPTABLE')).OnExit   := DateExit;
  THDBEdit(GetControl('TE_MONTANT'      )).OnChange := MontantChange;
  THDBEdit(GetControl('TE_MONTANTDEV'   )).OnChange := MontantChange;
  THDBEdit(GetControl('TE_MONTANT'      )).OnEnter  := MontantEnter;
  THDBEdit(GetControl('TE_MONTANTDEV'   )).OnEnter  := MontantEnter;
  {$ENDIF}
  NbDec := CalcDecimaleDevise(GetField('TE_DEVISE'));
  Taux := RetPariteEuro(GetField('TE_DEVISE'), StrToDateTime(GetControlText('TE_DATECOMPTABLE')));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  ObjCom : TObjCommissionTob; {24/03/05}
  T      : TOB;
  aDate  : TDateTime;
begin
  inherited;
  {Si on a modifier la date de valeur ou le montant, on recalcul les soldes}
  if CanModif then begin
    if IsFieldModified('TE_MONTANT') or IsFieldModified('TE_DATECOMPTABLE') then begin
      {13/10/06 : Pour le moment les commissions ne sont pas gérer su les comptes courants}
      if not CptCourant then begin
        BeginTrans;
        try
          {Mise à jour des commissions}
          T := TOB.Create('TRECRITURE', nil, -1);

          ObjCom := TObjCommissionTob.Create(GetField('TE_GENERAL'), GetField('TE_CODEFLUX'), GetField('TE_DATECOMPTABLE'));
          try
            {24/03/05 : reprise des valeurs au chargement de l'enregistrement, puuis des champs qui ont pu changer}
            T.Dupliquer(BufferAvantModif, True, True);
            T.SetDateTime('TE_DATEVALEUR', GetField('TE_DATEVALEUR'));
            T.SetDateTime('TE_DATECOMPTABLE', GetField('TE_DATECOMPTABLE'));
            T.SetDouble('TE_MONTANT', Arrondi(Valeur(GetField('TE_MONTANT')), v_pgi.OkDecV));
            T.SetDouble('TE_MONTANTDEV', Arrondi(Valeur(GetField('TE_MONTANTDEV')), NbDec));
            T.SetString('TE_CLEOPERATION', VarToStr(GetField('TE_CLEOPERATION')));
            T.SetString('TE_CLEVALEUR', VarToStr(GetField('TE_CLEVALEUR')));
            {24/03/05 : Utilisation de l'objet Tob plutôt que celui s'appuyant sur la structure TrEcriture}
            ObjCom.GenererCommissions(T);
            ObjCom.TobEcriture.UpdateDB;
          finally
            if Assigned(ObjCom) then FreeAndNil(ObjCom);
            if Assigned(T)      then FreeAndNil(T);
          end;
          CommitTrans;
        except
          on E : Exception do
            RollBack;
        end;
      end;
      aDate := Min(GetField('TE_DATECOMPTABLE'), BufferAvantModif.GetDateTime('TE_DATECOMPTABLE'));
      RecalculSolde(GetField('TE_GENERAL'), DateToStr(aDate), 0, True);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.OnLoadRecord;
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  v : THValComboBox;
  {$ELSE}
  v : THDBValComboBox;
  {$ENDIF}
begin
  inherited;
  {Ce cas n'est pas géré dans le OnArgument}
  if (GetField('TE_QUALIFORIGINE') = QUALIFCOMPTA)
     {JP 15/04/04 : Maintenant, il peut y avaoir des écritures comptables non réalisées :
                   and (GetField('TE_NATURE') = na_Realise)} then begin
    {$IFDEF EAGLCLIENT}
    v := THValComboBox(GetControl('TE_CODEFLUX'));
    {$ELSE}
    v := THDBValComboBox(GetControl('TE_CODEFLUX'));
    {$ENDIF}
    v.DataType := 'TZRUBRIQUE';
    v.Plus := 'RB_NATRUB = "TRE" AND RB_CLASSERUB = "TRE"';
  end;

  AssignDrapeau(TImage(GetControl('IDEV' )), GetField('TE_DEVISE'));
  AssignDrapeau(TImage(GetControl('IDEV1')), GetField('TE_DEVISE'));
  AssignDrapeau(TImage(GetControl('IDEV2')), GetField('TE_DEVISE'));
//  AssignDrapeau(TImage(GetControl('IDEV3')), GetField('TE_DEVISE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  Num : Integer;
begin
  inherited;
  {24/03/05 : Si on modifie les dates, il est bien de modifier les clefs}
  if IsFieldModified('TE_DATECOMPTABLE') then begin
    {20/11/06 : si l'on change de millésime : problème avec les soldes de réinitialisation}
    if TestDateEtMillesime(GetFieldAvantModif('TE_DATECOMPTABLE'), GetField('TE_DATECOMPTABLE'), True) or
       TestDateEtMillesime(GetFieldAvantModif('TE_DATEVALEUR'), GetField('TE_DATEVALEUR'), True) then begin
      LastError := 1;
      Exit;
    end;
    Num := StrToInt(Copy(VarToStr(GetField('TE_CLEOPERATION')), 7, 7));
    SetField('TE_CLEOPERATION', RetourneCleEcriture(VarToDateTime(GetField('TE_DATECOMPTABLE')), Num));
    SetField('TE_CLEVALEUR', RetourneCleEcriture(VarToDateTime(GetField('TE_DATEVALEUR')), Num));
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.DateExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {13/10/06 : sur les comptes courants, on ne peut modifier que les dates et la devise est la devise pivot}
  if not AfterInserting and CanModif and not CptCourant then begin
    if GetFieldAvantModif('TE_DATECOMPTABLE') <> StrToDateTime(GetControlText('TE_DATECOMPTABLE')) then begin
      {Récupération du taux de change du jour}
      Taux := RetPariteEuro(GetField('TE_DEVISE'), StrToDateTime(GetControlText('TE_DATECOMPTABLE')));
      if Taux = 0 then Taux := 1;  
      SetField('TE_COTATION', FormateMontant(1 / Taux, NBDECIMALTAUX));
      {Mise à jour du montant en Euro.
       21/10/04 : cf commentaires dans la procedure MontantChange}
      SetField('TE_MONTANT', Arrondi(Valeur(GetControlText('TE_MONTANTDEV')) * Taux, V_PGI.OkDecV));
      {Mise à jour de la date de valeur}
      if GetField('TE_NATURE') = na_Simulation then
        SetField('TE_DATEVALEUR', StrToDateTime(GetControlText('TE_DATECOMPTABLE')))
      else
        SetField('TE_DATEVALEUR', CalcDateValeur(GetField('TE_CODECIB'), GetField('TE_GENERAL'), StrToDateTime(GetControlText('TE_DATECOMPTABLE'))));
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.MontantChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {21/10/04 : Le problème est le même que partout ailleurs dans la Trésorerie : on part du principe que
              MontantEur * Cotation = MontantDev dans la Table TRECRITURE, cependant, "historiquement",
              RetPariteEuro renvoie le taux réel (1/Cotation). Il s'agit ici d'un cas où j'avais oublié
              d'uniformiser la gestion des devises}
  if not AfterInserting and CanModif and not CptCourant then begin
    {Pour éviter que les 2 Edits s'apellent mutuellement}
    if UpperCase(TComponent(Sender).Name) <> Composant then Exit;
    {Conversion du montant saisi}
    if UpperCase(TComponent(Sender).Name) = 'TE_MONTANT' then
      SetField('TE_MONTANTDEV', Arrondi(Valeur(GetControlText('TE_MONTANT')) / Taux, NbDec))
    else
      SetField('TE_MONTANT', Arrondi(Valeur(GetControlText('TE_MONTANTDEV')) * Taux, V_PGI.OkDecV));
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.MontantEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Composant := UpperCase(TComponent(Sender).Name);
end;

{01/08/07 : Nouvelle gestion des soldes : sur les réinitialisations on affiche les soldes ; Pour
            les autres flux, on se contente des Montant.
{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if F.FieldName = 'TE_CODEFLUX' then MajAffichage;
end;

{01/08/07 : Nouvelle gestion des soldes : sur les réinitialisations on affiche les soldes ; Pour
            les autres flux, on se contente des Montant.
{---------------------------------------------------------------------------------------}
procedure TOM_TRECRITURE.MajAffichage;
{---------------------------------------------------------------------------------------}
var
  InitOk : Boolean;
begin
  InitOk := VarToStr(GetField('TE_CODEFLUX')) = CODEREGULARIS;
  SetControlVisible('TTE_SOLDEDEV'      , InitOk);
  SetControlVisible('TTE_SOLDEDEVVALEUR', InitOk);
  SetControlVisible('TE_SOLDEDEVVALEUR' , InitOk);
  SetControlVisible('TE_SOLDEDEV'       , InitOk);
  SetControlVisible('IDEV1'             , InitOk);
  SetControlVisible('TTE_MONTANT'       , not InitOk);
  SetControlVisible('TTE_MONTANTDEV'    , not InitOk);
  SetControlVisible('TE_MONTANTDEV'     , not InitOk);
  SetControlVisible('TE_MONTANT'        , not InitOk);
end;

initialization
  RegisterClasses([TOM_TRECRITURE]);

end.
