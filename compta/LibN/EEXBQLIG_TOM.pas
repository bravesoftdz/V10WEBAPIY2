{-------------------------------------------------------------------------------------
    Version   |   Date | Qui | Commentaires
--------------------------------------------------------------------------------------
 8.01.001.001  08/12/07  JP   Création de l'unité : fiche de la table EEXBQLIG
 8.01.001.018  01/06/07  JP   FQ TRESO 10471 et COMPTA 20522 : refonte de la gestion du concept
                              de droit d'écriture dans EEXBQLIG
08.00.001.025  12/07/07  JP   FQ 21054 : Lors de la création d'un mouvement bancaire, on ne le pointe pas par défaut
08.00.001.025  17/07/07  JP   FQ 21099 : La date d'opération ne peut mordre sur une session fermée
08.00.001.025  18/07/07  JP   FQ 21126 : on interdit de créer un mouvement sur une session de rapprochement
08.00.001.025  19/07/07  JP   FQ 21113 : Initialisation des champs si création en série depuis le pointage
                              => Gestion d'une variable d'appel pour connaître la fiche appelante
08.00.002.002  30/07/07  JP   FQ 21193 : le contrôle sur la date d'opération lors de la création d'un mouvement est trop strict
08.00.002.002  30/07/07  JP   FQ 21194 : On ne peut supprimer un mouvement pointé
08.10.001.010  19/09/07  JP   FQ 21317 : Gestion des mouvements saisis manuellement lors de la suppression
--------------------------------------------------------------------------------------}
unit EEXBQLIG_TOM;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
   db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  Variants, UTob, Controls, Classes, SysUtils, HCtrls, HEnt1, UTOM;

type
  TOM_EEXBQLIG = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnArgument   (S : string); override;
    procedure OnDeleteRecord           ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnLoadRecord             ; override;
  private
    Compte       : string;
    NumReleve    : Integer;
    DateReleve   : TDateTime;
    LastDtReleve : TDateTime; {FQ 21193} 
    LaTob        : TOB;
    InChargement : Boolean;
    FChrAppelant : Char;

    procedure ChangementCompte;
    procedure MajNumeros;
    procedure MajBAcces;
    procedure ChangementDate(ChpDate : string);
    procedure BAccesOnClick (Sender : TObject);
  end;

procedure CPLanceFiche_EEXBQLIG(Range, Lequel, Arguments: string);

implementation

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  HDB,
  {$ENDIF EAGLCLIENT}
  ComCtrls, AglInit, Commun, uTOMEEXBQ, HTB97, ULibPointage, Ent1, HMsgBox, UProcGen,
  Constantes;

{Arguments : Action;APPEL;Compte;DateOpe;NumReleve;
APPEL : P -> pour Pointage , M -> pour le Mul
{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_EEXBQLIG(Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPEEXBQLIG', Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  cPlus : string;
begin
  inherited;
  NumReleve  := -1;
  DateReleve   := iDate1900; {30/07/07 : FQ 21193}
  LastDtReleve := iDate1900; {30/07/07 : FQ 21193}

  Ecran.HelpContext := 150;
  LaTob := TheTob;
  TheTob := nil;

  {Acces à la session de pointage}
  (GetControl('BACCES') as TToolbarButton97).OnClick := BAccesOnClick;

  {CEL_GENERAL contient des valeurs différentes selon le type de pointage}
  if EstPointageSurTreso then
    SetControlProperty('CEL_GENERAL', 'DATATYPE', 'TRBANQUECP')
  else if VH^.PointageJal then begin
    SetControlProperty('CEL_GENERAL', 'DATATYPE', 'TZJBANQUE');
    SetControlCaption('TCEL_GENERAL', TraduireMemoire('Journal'));
  end
  else
    SetControlProperty('CEL_GENERAL', 'DATATYPE', 'TTBANQUECP');
  {$IFDEF EAGLCLIENT}
  cPlus := FiltreBanqueCp(THEdit(GetControl('CEL_GENERAL')).DataType, tcb_Bancaire, '');
  THEdit(GetControl('CEL_GENERAL')).Plus := cPlus;
  {$ELSE}
  cPlus := FiltreBanqueCp(THDBEdit(GetControl('CEL_GENERAL')).DataType, tcb_Bancaire, '');
  THDBEdit(GetControl('CEL_GENERAL')).Plus := cPlus;
  {$ENDIF EAGLCLIENT}

  {Suppression de l'action}
  ReadTokenSt(s);
  {19/07/07 : FQ 21113 : Mémorisation de la fiche d'appel}
  FChrAppelant := StrToChr(ReadTokenSt(S));
  {01/06/07 : on regarde si l'on vient du pointage, pour cacher le bouton d'accès aux sessions de pointage}
  if FChrAppelant = FO_POINTAGE then
    (GetControl('BACCES') as TToolbarButton97).Enabled := False;

  if s <> '' then
    Compte := ReadTokenSt(S);
  if s <> '' then
    NumReleve := StrToInt(ReadTokenSt(S));
  if s <> '' then
    DateReleve := StrToDate(ReadTokenSt(S));
  {01/06/07 : FQ10471 et 20522}
  if not ExJaiLeDroitConcept(TConcept(ccSaisMvtBqe), False) then begin
    SetControlVisible('BINSERT', False);
    SetControlVisible('BDELETE', False);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.OnNewRecord;
{---------------------------------------------------------------------------------------}
var
  Sens : string;
  Q    : TQuery;
begin
  inherited;
  SetField('CEL_GENERAL', Compte);
  {Pour dire qu'il s'agit d'une ligne créée manuellement}
  SetField('CEL_VALIDE' , 'X');
  if Assigned(LaTob) then begin
    {Création d'un mouvement bancaire à partir d'une écriture sélectionnée dans le pointage}
    SetField('CEL_DATEOPERATION', LaTob.GetString('CLE_DATECOMPTABLE'));
    SetField('CEL_DATEVALEUR'   , LaTob.GetString('CLE_DATEVALEUR'));
    SetField('CEL_LIBELLE'      , LaTob.GetString('CLE_LIBELLE'));
    SetField('CEL_LIBELLE1'     , LaTob.GetString('CLE_REFINTERNE'));
    SetField('CEL_LIBELLE2'     , LaTob.GetString('CLE_REFEXTERNE'));
    SetField('CEL_LIBELLE3'     , LaTob.GetString('CLE_REFLIBRE'));
    SetField('CEL_REFPIECE'     , LaTob.GetString('CLE_NUMTRAITECHQ'));
    {Les montants des mouvements bancaires sont dans le même sens par rapport aux écritures
     Debit = Debit, ce qui revient à dire qu'ils sont inversés}
    SetField('CEL_DEBITEURO'    , LaTob.GetDouble('CLE_DEBIT'));
    SetField('CEL_CREDITEURO'   , LaTob.GetDouble('CLE_CREDIT'));
    SetField('CEL_CREDITDEV'    , LaTob.GetDouble('CLE_CREDITDEV'));
    SetField('CEL_DEBITDEV'     , LaTob.GetDouble('CLE_DEBITDEV'));

    if LaTob.GetString('CLE_CIB') <> '' then begin
      {Récupération du CIB par rapport au mode de paiement et au compte et au sens BANCAIRE}
      if (LaTob.GetDouble('CLE_MONTANT') >= 0) then Sens := 'ENC'
                                               else Sens := 'DEC';
      if EstPointageSurTreso then
        Sens := GetCodeCIB(Sens, LaTob.GetString('CLE_CIB'), Compte)
      else if VH^.PointageJal then begin
        Q := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL = "' + Compte + '"', True);
        try
          Sens := GetCodeCIB(Sens, LaTob.GetString('CLE_CIB'), Q.FindField('J_CONTREPARTIE').AsString, V_PGI.NoDossier);
        finally
          Ferme(Q);
        end;
      end
      else
        Sens := GetCodeCIB(Sens, LaTob.GetString('CLE_CIB'), Compte, V_PGI.NoDossier);
      if (Sens <> '') and (Sens <> '@@@') then SetField('CEL_CODEAFB',  Sens)
                                          else SetField('CEL_CODEAFB',  '');
    end
    else
      SetField('CEL_CODEAFB',  '');
  end
  else begin
    {30/07/07 : FQ 21193 : la DateReleve est maintenant toujours rattachée à un relevé. Sa valeur par défaut (iDate1900)
                sert à savoir s'il s'agit d'une création "brut", c'est-à-dire que l'information n'est pas connue lors de
                la création de l'enregistrement}
    if DateReleve > iDate1900 then begin
      SetField('CEL_DATEOPERATION', Date);
      SetField('CEL_DATEVALEUR'   , Date);
    end
    else begin
      SetField('CEL_DATEOPERATION', DateReleve);
      SetField('CEL_DATEVALEUR'   , DateReleve);
    end;
  end;

  ChangementDate('CEL_DATEOPERATION');//MajNumeros;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  {30/07/07 : FQ 21194 : On ne peut supprimer un mouvement pointé}
  if Trim(VarToStr(GetField('CEL_REFPOINTAGE'))) <> '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Il n''est pas possible de supprimer un mouvement pointé.');
  end;

  {19/09/07 : FQ 21317 : Gestion des mouvements saisis manuellement s'ils appartiennent à une session équilibrée}
  if (VarToStr(GetField('CEL_VALIDE')) = 'X') then begin
    if ExisteSQL('SELECT EE_AVANCEMENT FROM EEXBQ WHERE EE_GENERAL = "' + VarToStr(GetField('CEL_GENERAL')) +
                   '" AND EE_DATEPOINTAGE = "' + UsDateTime(DateReleve) + '" AND EE_AVANCEMENT = "X"') then begin
      if HShowMessage('0;' + Ecran.Caption + ';' +
                  TraduireMemoire('ATTENTION ! le mouvement (saisi manuellement) appartient à une amplitude de pointage terminée.') + #13#13 +
                  TraduireMemoire('Êtes-vous certain de vouloir le supprimer ?') + ';W;YN;N;N;', '', '') = mrNo then begin
        LastError := 1;
        LastErrorMsg := '';
      end;
    end;
  end;

  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if F.FieldName = 'CEL_GENERAL' then ChangementCompte
  else if F.FieldName = 'CEL_NUMRELEVE' then MajBAcces
  {30/07/07 : FQ 21193 : la gestion sur la date de valeur en Tréso me semble être une grosse erreur
  else if not (ctxTreso in V_PGI.PGIContexte) and (F.FieldName = 'CEL_DATEOPERATION') then ChangementDate('CEL_DATEOPERATION')
  else if (ctxTreso in V_PGI.PGIContexte) and (F.FieldName = 'CEL_DATEVALEUR') then ChangementDate('CEL_DATEVALEUR');}
  else if (F.FieldName = 'CEL_DATEOPERATION') then ChangementDate('CEL_DATEOPERATION');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  InChargement := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  B : Boolean;
  Q : TQuery;
begin
  {01/06/07 : FQ10471 et 20522 : si l'on est sur une opération issu d'un relevé, double message d'avertissement
              des dangers à modifier l'enregistrement}
  if VarToStr(GetField('CEL_VALIDE')) <> 'X' then begin
    B := PGIAsk(TraduireMemoire('Il est dangereux de modifier une ligne de relevé bancaire.') + #13 +
                TraduireMemoire('Souhaitez-vous poursuivre ?')) = mrYes;
    if B then
      B := PGIAsk(TraduireMemoire('Modifier une ligne de relevé bancaire risque de fausser vos soldes.') + #13 +
                TraduireMemoire('Souhaitez-vous Abandonner ?')) = mrNo;
    if not B then begin
      (GetControl('BDEFAIRE') as TToolbarButton97).Click;
      LastError := 1;
    end;
  end;

  inherited;
  if Trim(VarToStr(GetField('CEL_REFPOINTAGE'))) <> '' then begin
    (GetControl('BDEFAIRE') as TToolbarButton97).Click;
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Il n''est pas possible de modifier un mouvement pointé.');
  end
  else if Trim(VarToStr(GetField('CEL_CODEAFB'))) = '' then begin
    LastError := 2;
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le CIB');
    SetFocusControl('CEL_CODEAFB');
  end
  else if Trim(VarToStr(GetField('CEL_DEVISE'))) = '' then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner la devise');
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0;
    SetFocusControl('CEL_DEVISE');
  end

  {17/07/07 : FQ 21099 : je m'assure qu'il existe une référence de pointage ouverte pour la date d'opération
   18/07/07 : FQ 21126 : s'il s'agit d'une session de rapprochement, on interdit la création de mouvement}
  else begin
    {30/07/07 FQ 21193 : le contrôle sur la date est trop strict : la date d'opération doit être comprise entre
              la dernière session de pointage et la date de la session en cours}
    Q := OpenSQL('SELECT EE_AVANCEMENT, EE_ORIGINERELEVE FROM EEXBQ WHERE EE_GENERAL = "' + VarToStr(GetField('CEL_GENERAL'))
               + '" AND EE_DATEPOINTAGE = "' + UsDateTime(DateReleve) + '"', True);
    try
      if Q.EOF or (not Between(VarToDateTime(GetField('CEL_DATEOPERATION')), LastDtReleve, DateReleve)) then begin
        LastError := 2;
        LastErrorMsg := TraduireMemoire('La date d''opération ne correspond à aucune session de pointage ouverte.');
        SetFocusControl('CEL_DATEOPERATION');
      end
      else if Q.FindField('EE_AVANCEMENT').AsString = 'X' then begin
        LastError := 2;
        LastErrorMsg := TraduireMemoire('La date d''opération ne correspond à aucune session de pointage ouverte.');
        SetFocusControl('CEL_DATEOPERATION');
      end
      else if (DS.State = dsInsert) and IsReleveAuto(Q.FindField('EE_ORIGINERELEVE').AsString)
              {((Q.FindField('EE_ORIGINERELEVE').AsString = CODENEWPOINTAGE) or
               (Q.FindField('EE_ORIGINERELEVE').AsString = ORIGINERELEVE))} then begin
        LastError := 2;
        LastErrorMsg := TraduireMemoire('Il n''est pas possible de créer un mouvement sur une session de rapprochement');
        SetFocusControl('CEL_DATEOPERATION');
      end;
    finally
      Ferme(Q);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  inherited;
  if Assigned(LaTob) then begin
    T := TOB.Create('', LaTob, -1);
    AddChampPointage(T);
    T.SetString('CLE_JOURNAL'        , LIGNEBANCAIRE);
    T.SetString('CLE_GENERAL'        , VarToStr(GetField('CEL_GENERAL')));
    T.SetString('CLE_LIBELLE'        , VarToStr(GetField('CEL_LIBELLE')));
    T.SetString('CLE_REFINTERNE'     , VarToStr(GetField('CEL_LIBELLE1')));
    T.SetString('CLE_REFEXTERNE'     , VarToStr(GetField('CEL_LIBELLE2')));
    T.SetString('CLE_REFLIBRE'       , VarToStr(GetField('CEL_LIBELLE3')));
    T.SetString('CLE_NUMTRAITECHQ'   , VarToStr(GetField('CEL_REFPIECE')));
    T.SetString('CLE_CIB'            , VarToStr(GetField('CEL_CODEAFB')));
    T.SetString('CLE_DEVISE'         , VarToStr(GetField('CEL_DEVISE')));
    T.SetDouble('CLE_MONTANT'        , Valeur(VarToStr(GetField('CEL_CREDITEURO'))) - Valeur(VarToStr(GetField('CEL_DEBITEURO'))));
    T.SetDouble('CLE_MONTANTDEV'     , Valeur(VarToStr(GetField('CEL_CREDITDEV'))) - Valeur(VarToStr(GetField('CEL_DEBITDEV'))));
    T.SetDouble('CLE_CREDIT'         , Valeur(VarToStr(GetField('CEL_CREDITEURO'))));
    T.SetDouble('CLE_DEBIT'          , Valeur(VarToStr(GetField('CEL_DEBITEURO'))));
    T.SetDouble('CLE_DEBITDEV'       , Valeur(VarToStr(GetField('CEL_DEBITDEV'))));
    T.SetDouble('CLE_CREDITDEV'      , Valeur(VarToStr(GetField('CEL_CREDITDEV'))));
    T.SetInteger('CLE_NUMEROPIECE'   , ValeurI(VarToStr(GetField('CEL_NUMRELEVE'))));
    T.SetInteger('CLE_NUMLIGNE'      , ValeurI(VarToStr(GetField('CEL_NUMLIGNE'))));
    T.SetDateTime('CLE_DATECOMPTABLE', VarToDateTime(GetField('CEL_DATEOPERATION')));
    T.SetDateTime('CLE_DATEVALEUR'   , VarToDateTime(GetField('CEL_DATEVALEUR')));

    T.PutValue('CLE_REFPOINTAGE'  , '');
    T.PutValue('CLE_EXERCICE'     , '');
    T.PutValue('CLE_OLDREF'       , '');
    T.PutValue('CLE_DATEPOINTAGE' , iDate1900);
    T.PutValue('CLE_NUMECHE'      , 0);
    T.PutValue('CLE_QUALIFPIECE'  , VarToStr(GetField('CEL_REFPIECE')));
    T.PutValue('CLE_MANUEL'       , 'X');
    T.PutValue('CLE_DATEECHEANCE' , Now);
    {12/07/07 : FQ 21054 : On ne pointe pas l'éciture par défaut}
    T.PutValue('CLE_POINTE'       , '');
    T.PutValue('MODIFIE'          , 'X');
  end;
  TheTob := LaTob;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.ChangementCompte;
{---------------------------------------------------------------------------------------}
var
  Q    : TQuery;
  Plus : string;
  lCpt : string;
begin
  lCpt := VarToStr(GetField('CEL_GENERAL'));
  {Filtre sur les CIB}
  if not (CtxPcl in V_PGI.PGIContexte) then begin
    if VH^.PointageJal then
      Plus := '(SELECT BQ_BANQUE FROM BANQUECP, JOURNAL WHERE J_CONTREPARTIE = BQ_GENERAL AND J_JOURNAL = "' + lCpt + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '")'
    else if EstPointageSurTreso then
      Plus := '(SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_CODE = "' + lCpt + '")'
    else
      Plus := '(SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_GENERAL = "' + lCpt + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '")';
    {$IFDEF EAGLCLIENT}
    (GetControl('CEL_CODEAFB') as THEdit).Plus := 'TCI_BANQUE = ' + Plus;
    {$ELSE}
    (GetControl('CEL_CODEAFB') as THDBEdit).Plus := 'TCI_BANQUE = ' + Plus;
    {$ENDIF EAGLCLIENT}
  end;

  if VH^.PointageJal then
    Q := OpenSQL('SELECT BQ_DEVISE, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CODEIBAN FROM BANQUECP, JOURNAL WHERE ' +
                 'BQ_GENERAL = J_CONTREPARTIE AND J_JOURNAL = "' + lCpt + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"', True)
  else if EstPointageSurTreso then
    Q := OpenSQL('SELECT BQ_DEVISE, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CODEIBAN FROM BANQUECP WHERE ' +
                 'BQ_CODE = "' + lCpt + '"', True)
  else
    Q := OpenSQL('SELECT BQ_DEVISE, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CODEIBAN FROM BANQUECP WHERE ' +
                 'BQ_GENERAL = "' + lCpt + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"', True);
  try
    if not Q.EOF then begin
      if Q.FindField('BQ_DEVISE').AsString <> V_PGI.DevisePivot then begin
        SetControlVisible('CEL_DEBITEURO', False);
        SetControlVisible('CEL_CREDITEURO', False);
        SetControlVisible('CEL_DEBITDEV', True);
        SetControlVisible('CEL_CREDITDEV', True);
        SetControlCaption('CDEV', Q.FindField('BQ_DEVISE').AsString);
        SetControlCaption('DDEV', Q.FindField('BQ_DEVISE').AsString);
      end
      else begin
        SetControlVisible('CEL_DEBITEURO', True);
        SetControlVisible('CEL_CREDITEURO', True);
        SetControlVisible('CEL_DEBITDEV', False);
        SetControlVisible('CEL_CREDITDEV', False);
        SetControlCaption('CDEV', V_PGI.DevisePivot);
        SetControlCaption('DDEV', V_PGI.DevisePivot);
      end;

      if DS.State in [dsInsert, dsEdit] then begin
        SetField('CEL_DEVISE', Q.FindField('BQ_DEVISE').AsString);
        if Trim(Q.FindField('BQ_ETABBQ').AsString) = '' then
          SetField('CEL_RIB', Q.FindField('BQ_CODEIBAN').AsString)
        else
          SetField('CEL_RIB', Q.FindField('BQ_ETABBQ').AsString + '-' +
                              Q.FindField('BQ_GUICHET').AsString + '-' +
                              Q.FindField('BQ_NUMEROCOMPTE').AsString);
      end;
    end
    else begin
      SetControlVisible('CEL_DEBITDEV', False);
      SetControlVisible('CEL_CREDITDEV', False);
      SetControlCaption('CDEV', V_PGI.DevisePivot);
      SetControlCaption('DDEV', V_PGI.DevisePivot);
    end;
  finally
    Ferme(Q);
  end;

  if lCpt <> Compte then begin
    ChangementDate('CEL_DATEOPERATION');//MajNumeros;
    Compte := lCpt;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.MajNumeros;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  lCpt : string;
begin
  inherited;
  lCpt := VarToStr(GetField('CEL_GENERAL'));
  (*
  {On calcule le numéro du relevé s'il n'est pas initialisé ou bien si le général a changé}
  if (NumReleve = -1) or (Compte <> lCpt) then begin
    Q := OpenSQL('SELECT MAX(EE_NUMRELEVE) FROM EEXBQ WHERE EE_GENERAL = "' + lCpt +
                 '" AND EE_DATEPOINTAGE <= "' + UsDateTime(DateReleve) + '"', True);
    if not Q.EOF then
      NumReleve := Q.Fields[0].AsInteger;
    Ferme(Q);
  end;*)

  SetField('CEL_NUMRELEVE', NumReleve);

  Q := OpenSQL('SELECT MAX(CEL_NUMLIGNE) FROM EEXBQLIG WHERE CEL_GENERAL = "' + lCpt + '" AND CEL_NUMRELEVE = ' + IntToStr(NumReleve), True);
  if not Q.EOF then
    SetField('CEL_NUMLIGNE', Q.Fields[0].AsInteger + 1)
  else
    SetField('CEL_NUMLIGNE', 1);
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.ChangementDate(ChpDate : string);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if InChargement and (DS.State = dsEdit) then
    InChargement := False;
  if DS.State in [dsInsert, dsEdit] then begin
    if (not Between(VarToDateTime(GetField(ChpDate)), LastDtReleve, DateReleve)) or
       (LastDtReleve = iDate1900) or (DateReleve = iDate1900) then begin
      {30/07/07 : FQ 21193 : recherche de la date du relevé auquel va appartenir le mouvement}
      Q := OpenSQL('SELECT MIN(EE_DATEPOINTAGE), MAX(EE_NUMRELEVE) FROM EEXBQ WHERE EE_GENERAL = "' + VarToStr(GetField('CEL_GENERAL')) +
                   '" AND EE_DATEPOINTAGE >= "' + UsDateTime(VarToDateTime(GetField(ChpDate))) + '"', True);
      if not Q.EOF then begin
        DateReleve := Q.Fields[0].AsDateTime;
        NumReleve  := Q.Fields[1].AsInteger;
      end;
      Ferme(Q);

      {30/07/07 : FQ 21193 : recherche de la date du précédent relevé}
      Q := OpenSQL('SELECT MAX(EE_DATEPOINTAGE) FROM EEXBQ WHERE EE_GENERAL = "' + VarToStr(GetField('CEL_GENERAL')) +
                   '" AND EE_DATEPOINTAGE < "' + UsDateTime(VarToDateTime(GetField(ChpDate))) + '"', True);
      if not Q.EOF then LastDtReleve := Q.Fields[0].AsDateTime;
      Ferme(Q);
    end;

    MajNumeros;
  end
  else begin
    DateReleve := VarToDateTime(GetField('CEL_DATEOPERATION'));
    NumReleve  := ValeurI(VarToStr(GetField('CEL_NUMRELEVE')));
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.BAccesOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CPLanceFiche_PointageFic(GetField('CEL_GENERAL') + ';' + DateToStr(GetField('CEL_DATEPOINTAGE')) + ';' + GetField('CEL_REFPOINTAGE') + ';' + IntToStr(GetField('CEL_NUMRELEVE')), 'ACTION=CONSULTATION')
end;

{---------------------------------------------------------------------------------------}
procedure TOM_EEXBQLIG.MajBAcces;
{---------------------------------------------------------------------------------------}
begin
  {Pour accéder à la référence de pointage, il faut qu'elle existe ! Pour cela on s'assure que NUMRELEVE
   est supérieur à zéro et que l'on n'est pas en insertion}
  (GetControl('BACCES') as TToolbarButton97).Visible := not (DS.State in [dsInsert]) and (ValeurI(GetField('CEL_NUMRELEVE')) > 0);
end;

initialization
  RegisterClasses([TOM_EEXBQLIG]);

end.

