{-------------------------------------------------------------------------------------
  Version    | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
              06/12/01  BT   Création de l'unité
1.01.001.001  04/03/04  JP   Mise en place d'un contrôle dans le OnUpdateRecord et
                             le OnDeleteRecor sur le Code Flux pour interdire les
                             modifications sur les 3 flux EQD, EQR, REI fournis dans
                             la "maquette"
1.05.001.002  22/03/04  JP   Affinage de ce qui précède sur le OnUpdateRecord
6.00.020.002  23/11/04  JP   FQ 10132 : on cache les boutons Nouveau et Consultation
                             dans le LookupList pour éviter de tirer la comptabilité
6.30.001.004  15/03/05  JP   FQ 10221 : Gestion des flux de référence dans le OnUpdate
7.09.001.001  15/09/06  JP   Modification de la gestion de la TVA : elle est toujours sur décaissement !!
7.09.001.001  12/10/06  JP   Ajout de la classe ICC en mode multi sociétés
7.09.001.009  20/03/07  JP   Corrections concernant les ICC dans le OnChangeField
--------------------------------------------------------------------------------------}
unit TomFluxTreso ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DBCtrls, HDB,
  {$ENDIF}
  StdCtrls, Controls, Classes, forms, SysUtils, HCtrls, HEnt1, UTOM, UTob, HMsgBox;

type
  TOM_FLUXTRESO = class (TOM)
    procedure OnChangeField(F : TField); override;
    procedure OnArgument   (S : string); override;
    procedure OnClose                  ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnDeleteRecord           ; override;
    procedure OnNewRecord              ; override;
  protected
    TobRub : TOB;
    procedure GererTVAChange(Sender : TObject);
    procedure AppelGeneraux (Sender : TObject); {23/11/04 : FQ 10132}
    procedure CodeFluxTresoOnKeyPress(Sender: TObject; var Key: Char);
  end ;

procedure TRLanceFiche_FluxTreso(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

implementation

uses
  Commun, Constantes, LookUp;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_FluxTreso(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  inherited ;
  Ecran.HelpContext := 150;
  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('TFT_FLUX')).OnKeyPress := CodeFluxTresoOnKeyPress;
  TCheckBox(GetControl('TFT_ASSUJETTITVA')).OnClick :=  GererTVAChange;
  THEdit(GetControl('TFT_GENERAL')).OnElipsisClick := AppelGeneraux;{23/11/04 : FQ 10132}
  {12/10/06 : On exclut la Classe ICC si on n'est pas en MultiDossier}
  if not IsTresoMultiSoc then
    THValComboBox(GetControl('TFT_CLASSEFLUX')).Plus := 'AND CO_CODE <> "' + cla_ICC + '"'; 
  {$ELSE}
  THDBEdit(GetControl('TFT_FLUX')).OnKeyPress := CodeFluxTresoOnKeyPress;
  TDBCheckBox(GetControl('TFT_ASSUJETTITVA')).OnClick :=  GererTVAChange;
  THDBEdit(GetControl('TFT_GENERAL')).OnElipsisClick := AppelGeneraux;{23/11/04 : FQ 10132}
  {12/10/06 : On exclut la Classe ICC si on n'est pas en MultiDossier}
  if not IsTresoMultiSoc then
    THDBValComboBox(GetControl('TFT_CLASSEFLUX')).Plus := 'AND CO_CODE <> "' + cla_ICC + '"';
  {$ENDIF}
  TobRub := TOB.Create('', nil, -1);
  Q := OpenSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_NATRUB = "TRE" AND RB_CLASSERUB = "TRE"', True);
  TobRub.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.OnClose;
{---------------------------------------------------------------------------------------}
begin
  FreeAndNil(TobRub);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.CodeFluxTresoOnKeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ValidCodeOnKeyPress(Key);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited;
  {JP 10/09/03 : Les écritures synchronisées depuis la compta ont comme code de flux un code rubrique
                 On empêche donc de saisir un code flux qui corresponde à un code rubrique}
  if (F.FieldName = 'TFT_FLUX') and (DS.State in [dsInsert, dsEdit]) then begin
    T := TobRub.FindFirst(['RB_RUBRIQUE'], [F.AsString], False);
    if Assigned(T) then begin
      HShowMessage('0;Flux de trésorerie; Ce code est déjà utilisé dans les rubriques !;W;O;O;O;', '', '');
      THEdit(GetControl('TFT_FLUX')).SetFocus;
    end;
  end
  else if (F.FieldName = 'TFT_CLASSEFLUX') then begin
    {12/10/06 : A priori pas de TVA et de commission sur les intérêts de comptes courants}
    {A priori(!) les flux de commissions ne sont pas soumis à commissions de mouvement}
    SetControlEnabled('TFT_COMMOUVEMENT', (GetField('TFT_CLASSEFLUX') <> cla_Commission) and
                                          (GetField('TFT_CLASSEFLUX') <> cla_ICC));
    {15/09/06 : la TVA n'est que sur les commissions / Frais}
    SetControlEnabled('TFT_ASSUJETTITVA', GetField('TFT_CLASSEFLUX') = cla_Commission);
    SetControlEnabled('TFT_CODETVA', GetField('TFT_CLASSEFLUX') = cla_Commission);
    SetControlEnabled('TTFT_CODETVA', GetField('TFT_CLASSEFLUX') = cla_Commission);

    {18/10/06 : Pour les ICC, on force le code cib à 888
     20/03/07 : Test sur l'état du datasource et setField au lieu de SetControlText}
    if DS.State in [dsInsert, dsEdit] then
      if GetField('TFT_CLASSEFLUX') = cla_ICC then SetField('TFT_CODECIB', CODECIBCOURANT);
      
    SetControlEnabled('TFT_CODECIB', GetField('TFT_CLASSEFLUX') <> cla_ICC);

    {Pour éviter de redonner le focus au code flux}
    if not AfterInserting then SetFocusControl('TFT_CODECIB');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited ;
  {JP 04/03/04 : On interdit la modification des enregistrements fournis dans la maquette}
  if (BufferAvantModif.GetString('TFT_CLASSEFLUX') = cla_Reference) then begin
    {JP 22/03/04 : Ajout de ce test, car il faut pouvoir affecter le compte de virement interne (580XXXXX)}
    if (BufferAvantModif.GetString('TFT_FLUX') <> GetField('TFT_FLUX')) or
       (BufferAvantModif.GetString('TFT_TYPEFLUX') <> GetField('TFT_TYPEFLUX')) or
       (BufferAvantModif.GetString('TFT_CLASSEFLUX') <> GetField('TFT_CLASSEFLUX')) then begin
      {15/03/05 : FQ 10221 : On peut modifier si les deux conditions sont remplies : Superviseur et SAV}
      if (not V_PGI.SAV) or (not V_PGI.Superviseur) then begin
        LastErrorMsg := TraduireMemoire('Vous n''avez pas le droit de modifier l''identité de ce type de flux.');
        LastError := 1;
        Exit;
      end;
    end;
  end
  else if (GetField('TFT_CLASSEFLUX') = cla_Commission) then begin
    if ExisteSQL('SELECT TTL_TYPEFLUX FROM TYPEFLUX WHERE TTL_SENS = "C" AND TTL_TYPEFLUX = "' + GetField('TFT_TYPEFLUX') + '"') then begin
      LastErrorMsg := TraduireMemoire('Le type de flux pour une commission doit être débiteur.');
      LastError := 1;
      Exit;
    end
    else if GetField('TFT_COMMOUVEMENT') = 'X' then begin
      LastErrorMsg := TraduireMemoire('Les flux de commission ne peuvent pas être débiteurs.');
      LastError := 1;
      Exit;
    end;
  end
  {15/03/05 : FQ 10221 : On est en création}
  else if (BufferAvantModif.GetString('TFT_CLASSEFLUX') = '') and
          (VarToStr(GetField('TFT_CLASSEFLUX')) = cla_Reference) then begin
    {FQ 10221 : On peut modifier si les deux conditions sont remplies : Superviseur et SAV}
    if (not V_PGI.SAV) or (not V_PGI.Superviseur) then begin
      LastErrorMsg := TraduireMemoire('Vous n''avez pas le droit de créer un flux de classe référence.');
      LastError := 1;
      Exit;
    end;
  end;

  if GetField('TFT_GENERAL') = '' then begin
    LastErrorMsg := TraduireMemoire('Vous devez renseigner un compte général.');
    LastError := 1;
    SetFocusControl('TFT_GENERAL');
    Exit;
  end
  {23/04/04 : Tant que l'on ne gère pas les comptes auxiliaires dans la trésorerie}
  else if not LookupValueExist(GetControl('TFT_GENERAL')) then begin
    LastErrorMsg := TraduireMemoire('Le compte général ne doit pas être un compte collectif.');
    LastError := 1;
    SetFocusControl('TFT_GENERAL');
    Exit;
  end;

  if (GetField('TFT_ASSUJETTITVA') = 'X') and (Trim(GetField('TFT_CODETVA')) = '') then begin
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le code TVA.');
    LastError := 1;
    SetFocusControl('TFT_CODETVA');
    Exit;
  end;

  {03/06/04 : On force aussi la saisie d'un type de flux et d'un code cib}
  if (Trim(GetField('TFT_TYPEFLUX')) = '') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un type de flux.');
    SetFocusControl('TFT_TYPEFLUX');
    Exit;
  end
  else if (Trim(GetField('TFT_CODECIB')) = '') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un code cib.');
    SetFocusControl('TFT_CODECIB');
    Exit;
  end;

  {JP 10/09/03 : Les écritures synchronisées depuis la compta ont comme code de flux un code rubrique
                 On empêche donc de saisir un code flux qui corresponde à un code rubrique}
  T := TobRub.FindFirst(['RB_RUBRIQUE'], [GetField('TFT_FLUX')], False);
  if Assigned(T) then begin
    LastErrorMsg := TraduireMemoire('Ce code est déjà utilisé dans les rubriques.');
    LastError := 1;
    SetFocusControl('TFT_FLUX');
    Exit;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.OnDeleteRecord ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  {JP 04/03/04 : On interdit la modification des enregistrements fournis dans la maquette}
  if (GetField('TFT_CLASSEFLUX') = cla_Reference) then begin
    HShowMessage('0;' + Ecran.Caption + ';Vous n''avez pas le droit de supprimer ce type de flux;W;O;O;O', '', '');
    LastError := 1;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.GererTVAChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Gerer : Boolean;
begin
  Gerer := GetCheckBoxState('TFT_ASSUJETTITVA') = cbChecked;
  SetControlEnabled('TFT_CODETVA', Gerer);
  SetControlEnabled('TTFT_CODETVA', Gerer);
  SetControlEnabled('TFT_TVAENCAIS', Gerer);
end;

{FQ 10132 : Pour ne pas appeler toute la Compta, il faut interdire la loupe et la création,
            chose que je ne peux pas faire dans le DispatchTT => j'intercepte
{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.AppelGeneraux(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LookupList(GetControl('TFT_GENERAL'), 'Liste des généraux', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE',
             'G_COLLECTIF <> "X"', 'G_GENERAL', True, 0); {0 permet de cacher les deux boutons}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FLUXTRESO.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {27/04/05 : FQ 10244 : s'il y a un filtre sur la tablette TROPCVM au moment du DispatchTT,
              il est repris dans le champ de l'index + FQ AGL 11512}
  SetField('TFT_FLUX', '');
  {15/09/06 : la TVA est théoriquement que sur les frais bancaires, donc sur décaissement}
  SetField('TFT_TVAENCAIS', '-');
end;

initialization
  RegisterClasses([TOM_FLUXTRESO]);

end.

