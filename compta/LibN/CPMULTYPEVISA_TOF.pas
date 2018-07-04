{ Unité : Source TOF dMul : CPMULTYPEVISA ()
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  02/02/06   JP   Création de l'unité
 8.01.001.020  14/06/07   JP   Ajout du bouton de duplication
--------------------------------------------------------------------------------------}
unit CPMULTYPEVISA_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eMul,
  {$ELSE}
  FE_Main, db, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} Mul, HDB,
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  sysutils, UTob, uTOF, HCtrls;

  type
  TOF_CPMULTYPEVISA = class (TOF)
    procedure OnArgument(S : string); override;
  private
    procedure RecupControls;
    procedure MajAffichage (Value, Ind : string; CodeSender : Boolean = False);
    procedure MajTableLibre(Ind : string);
  public
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF EAGLCLIENT}
    TypeLibre1  : THValComboBox;
    TypeLibre2  : THValComboBox;
    TypeLibre3  : THValComboBox;
    CodeLibre1  : THValComboBox;
    CodeLibre2  : THValComboBox;
    CodeLibre3  : THValComboBox;
    TexteLibre1 : THEdit;
    TexteLibre2 : THEdit;
    TexteLibre3 : THEdit;

    procedure BInsertClick(Sender : TObject);
    procedure BDupliqClick(Sender : TObject); {JP 14/06/07}
    procedure BOuvrirClick(Sender : TObject);
    procedure BDeleteClick(Sender : TObject);
    procedure SlctAllClick(Sender : TObject);
    procedure ControlExit (Sender : TObject);
    procedure CodeOnExit  (Sender : TObject);
  end;

procedure CpLanceFiche_TypeVisaMul;

implementation

uses
  HEnt1, HMsgBox, Ent1, HTB97, CPTYPEVISA_TOM, ULibBonAPayer, Forms, UObjGen;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_TypeVisaMul;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMULTYPEVISA', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509010;

  FListe := TFMul(Ecran).FListe;

  TToolbarButton97(GetControl('BINSERT'   )).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BOUVRIR'   )).OnClick := BOuvrirClick;
  TToolbarButton97(GetControl('BDELETE'   )).OnClick := BDeleteClick;
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;
  TToolbarButton97(GetControl('BDUPLIQUER')).OnClick := BDupliqClick; {JP 14/06/07}

  FListe.OnDblClick := BOuvrirClick;
  
  RecupControls;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_TypeVisaFiche('', '', 'ACTION=CREATION');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.BOuvrirClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_TypeVisaFiche('', VarToStr(GetField('CTI_CODEVISA')), 'ACTION=MODIFICATION');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}

    {------------------------------------------------------------------------}
    procedure DetruireTypeVisa;
    {------------------------------------------------------------------------}
    var
      Code : string;
    begin
      Code := TFMul(Ecran).Q.FindField('CTI_CODEVISA').AsString;
      if CanDeleteTypeVisa(Code) then begin
        ExecuteSQL('DELETE FROM CPTYPEVISA WHERE CTI_CODEVISA = "' + Code + '"');
        Inc(p);
      end
      else
        HShowMessage('0;' + Ecran.Caption + ';Le type visa "' + Code  + '" est utilisé dans les bons à payer'#13 +
                     'et ne peut être supprimé.;I;O;O;O;', '', '');
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  p := 0;

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end
  else if HShowMessage('0;' + Ecran.Caption + ';Êtes vous sûr de vouloir supprimer la sélection ?;Q;YN;N;N;', '', '') <> mrYes then
    Exit;

  BeginTrans;
  try
    {$IFNDEF EAGLCLIENT}
    TFMul(Ecran).Q.First;
    if F.AllSelected then
      while not TFMul(Ecran).Q.EOF do begin
        DetruireTypeVisa;
        TFMul(Ecran).Q.Next;
      end
    else
    {$ENDIF}

    for n := 0 to F.nbSelected - 1 do begin
      F.GotoLeBookmark(n);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
      {$ENDIF}
      DetruireTypeVisa;
    end;
    CommitTrans;
  except
    on E : Exception do begin
      RollBack;
      HShowMessage('0;' + Ecran.Caption + '; Traitement interrompu :'#13 + E.Message + ';E;O;O;O;', '', '');
    end;
  end;

  if p > 0 then
    HShowMessage('0;' + Ecran.Caption + ';' + IntToStr(p) + ' enregistrement(s) a (ont) été supprimé(s).;I;O;O;O;', '', '');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.SlctAllClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.MajAffichage(Value, Ind : string; CodeSender : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  Code  : string;
  AxeOk : Boolean;
begin
  if Value = '' then begin
    SetControlEnabled('CTI_CODELIBRE'   + Ind, False);
    SetControlEnabled('TCTI_CODELIBRE'  + Ind, False);
    SetControlEnabled('TCTI_TEXTELIBRE' + Ind, False);
    SetControlEnabled('CTI_TEXTELIBRE'  + Ind, False);
    SetControlEnabled('TCTI_AXE'        + Ind, False);
    SetControlEnabled('CTI_AXE'         + Ind, False);
    SetControlProperty('CTI_TEXTELIBRE' + Ind, 'ELIPSISBUTTON', False);
    SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', '');
  end

  else begin
    SetControlEnabled('TCTI_TEXTELIBRE' + Ind, True);
    SetControlEnabled('CTI_TEXTELIBRE'  + Ind, True);

    if (Value = tyt_Devise) or (Value = tyt_RefInterne) then begin
      SetFocusControl('CTI_TEXTELIBRE'    + Ind);
      SetControlEnabled('CTI_CODELIBRE'   + Ind, False);
      SetControlEnabled('TCTI_CODELIBRE'  + Ind, False);
      SetControlEnabled('TCTI_AXE'        + Ind, False);
      SetControlEnabled('CTI_AXE'         + Ind, False);
      SetControlText('CTI_CODELIBRE'      + Ind, '');
      SetControlText('CTI_AXE'            + Ind, '');
      SetControlProperty('CTI_TEXTELIBRE' + Ind, 'ELIPSISBUTTON', False);
      SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', '');
    end

    else begin
      AxeOk := True;
      {Récupération éventuelle de l'indice de la table libre à traiter}
      Code := GetControlText('CTI_CODELIBRE' + Ind);
      if Value = tyt_TLSection then begin
        AxeOk := GetControlEnabled('CTI_AXE' + Ind);
        SetControlEnabled('CTI_AXE'  + Ind, True);
        SetControlEnabled('TCTI_AXE' + Ind, True);
      end
      else begin
        SetControlEnabled('TCTI_AXE' + Ind, False);
        SetControlEnabled('CTI_AXE'  + Ind, False);
      end;

      if not CodeSender then begin
        {Le sender est le type, on force l'ActiveControl sur le code}
        if not GetControlEnabled('CTI_CODELIBRE' + Ind) then begin
          SetControlEnabled('CTI_CODELIBRE'  + Ind, True);
          SetControlEnabled('TCTI_CODELIBRE' + Ind, True);
          SetFocusControl('CTI_CODELIBRE' + Ind);
        end;
      end
      else if Value = tyt_TLSection then begin
        {Pour ne pas faire le SetFocus si on ne passe pas au contrôle suivant, Maj + Tab par exemple}
        if not AxeOk then
          SetFocusControl('CTI_AXE' + Ind);
      end;

      if Length(Code) = 3  then Code := Code[3]
                           else Code := '0';

      SetControlEnabled('CTI_CODELIBRE'   + Ind, True);
      SetControlEnabled('TCTI_CODELIBRE'  + Ind, True);
      SetControlProperty('CTI_TEXTELIBRE' + Ind, 'ELIPSISBUTTON', True);

      if Value = tyt_TLSection then begin
        SetControlProperty('CTI_CODELIBRE'  + Ind, 'DATATYPE', 'TTTABLESLIBRESSEC');
        SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', 'TZNATSECT' + Code);
      end
      else begin
        if Value = tyt_TLAuxiliaire then begin
          SetControlProperty('CTI_CODELIBRE'  + Ind, 'DATATYPE', 'TTTABLESLIBRESAUX');
          SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', 'TZNATTIERS' + Code);
        end
        else if Value = tyt_TLEcriture then begin
          SetControlProperty('CTI_CODELIBRE'  + Ind, 'DATATYPE', 'TTTABLESLIBRESECR');
          SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', 'TZNATECR' + Code);
        end
        else if Value = tyt_TLGeneral then begin
          SetControlProperty('CTI_CODELIBRE'  + Ind, 'DATATYPE', 'TTTABLESLIBRESGEN');
          SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', 'TZNATGENE' + Code);
        end;
        SetControlText('CTI_AXE' + Ind, '');
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.RecupControls;
{---------------------------------------------------------------------------------------}
begin
  TypeLibre1  := THValComboBox(GetControl('CTI_TYPELIBRE1'));
  TypeLibre2  := THValComboBox(GetControl('CTI_TYPELIBRE2'));
  TypeLibre3  := THValComboBox(GetControl('CTI_TYPELIBRE3'));
  CodeLibre1  := THValComboBox(GetControl('CTI_CODELIBRE1'));
  CodeLibre2  := THValComboBox(GetControl('CTI_CODELIBRE2'));
  CodeLibre3  := THValComboBox(GetControl('CTI_CODELIBRE3'));
  TexteLibre1 := THEdit(GetControl('CTI_TEXTELIBRE1'));
  TexteLibre2 := THEdit(GetControl('CTI_TEXTELIBRE2'));
  TexteLibre3 := THEdit(GetControl('CTI_TEXTELIBRE3'));

  TypeLibre1.OnExit := ControlExit;
  TypeLibre2.OnExit := ControlExit;
  TypeLibre3.OnExit := ControlExit;

  CodeLibre1.OnExit := CodeOnExit;
  CodeLibre2.OnExit := CodeOnExit;
  CodeLibre3.OnExit := CodeOnExit;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.ControlExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  MajAffichage(THValComboBox(Sender).Value, THValComboBox(Sender).Name[Length(THValComboBox(Sender).Name)]);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.CodeOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  MajTableLibre(THValComboBox(Sender).Name[Length(THValComboBox(Sender).Name)]);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.MajTableLibre(Ind : string);
{---------------------------------------------------------------------------------------}
begin
  MajAffichage(THValComboBox(GetControl('CTI_TYPELIBRE' + Ind)).Value, Ind, True);
end;

{Ajout de l'option de duplication (demande de SIC) pour accélérer le paramétrage
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULTYPEVISA.BDupliqClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}
  TobVisa  : TOB;
  TobCode  : TOB;
  LastCode : string;


    {------------------------------------------------------------------------}
    function _DupliqueTypeVisa : Boolean;
    {------------------------------------------------------------------------}
    var
      Code : string;
      T : TOB;
      Q : TQuery;

        {-----------------------------------------------------------------}
        function _GenereCode : Boolean;
        {-----------------------------------------------------------------}
        var
          T : TOB;
        begin
          repeat
            TObjCodeCombo.GetCodeComboUnique(LastCode);
            T := TobCode.FindFirst(['CTI_CODEVISA'], [LastCode], True);
          until T = nil;
          Result := LastCode <> '$$$';
        end;


    begin
      Result := _GenereCode;
      if not Result then Exit;

      Code := TFMul(Ecran).Q.FindField('CTI_CODEVISA').AsString;
      Q := OpenSQL('SELECT * FROM CPTYPEVISA WHERE CTI_CODEVISA = "' + Code + '"', True);
      try
        T := TOB.Create('CPTYPEVISA', TobVisa, -1);
        T.PutValue('CTI_CODEVISA'     , LastCode);
        T.PutValue('CTI_LIBELLE'      , Q.FindField('CTI_LIBELLE'      ).AsString + ' (BIS)');
        T.PutValue('CTI_ABREGE'       , Q.FindField('CTI_ABREGE'       ).AsString);
        T.PutValue('CTI_NBVISA'       , Q.FindField('CTI_NBVISA'       ).AsInteger);
        T.PutValue('CTI_NATUREPIECE'  , Q.FindField('CTI_NATUREPIECE'  ).AsString);
        T.PutValue('CTI_MONTANTMIN'   , Q.FindField('CTI_MONTANTMIN'   ).AsFloat);
        T.PutValue('CTI_MONTANTMAX'   , Q.FindField('CTI_MONTANTMAX'   ).AsFloat);
        T.PutValue('CTI_ETABLISSEMENT', Q.FindField('CTI_ETABLISSEMENT').AsString);
        T.PutValue('CTI_TYPELIBRE1'   , Q.FindField('CTI_TYPELIBRE1'   ).AsString);
        T.PutValue('CTI_TYPELIBRE2'   , Q.FindField('CTI_TYPELIBRE2'   ).AsString);
        T.PutValue('CTI_TYPELIBRE3'   , Q.FindField('CTI_TYPELIBRE3'   ).AsString);
        T.PutValue('CTI_CODELIBRE1'   , Q.FindField('CTI_CODELIBRE1'   ).AsString);
        T.PutValue('CTI_CODELIBRE2'   , Q.FindField('CTI_CODELIBRE2'   ).AsString);
        T.PutValue('CTI_CODELIBRE3'   , Q.FindField('CTI_CODELIBRE3'   ).AsString);
        T.PutValue('CTI_TEXTELIBRE1'  , Q.FindField('CTI_TEXTELIBRE1'  ).AsString);
        T.PutValue('CTI_TEXTELIBRE2'  , Q.FindField('CTI_TEXTELIBRE2'  ).AsString);
        T.PutValue('CTI_TEXTELIBRE3'  , Q.FindField('CTI_TEXTELIBRE3'  ).AsString);
        T.PutValue('CTI_AXE1'         , Q.FindField('CTI_AXE1'         ).AsString);
        T.PutValue('CTI_AXE2'         , Q.FindField('CTI_AXE2'         ).AsString);
        T.PutValue('CTI_AXE3'         , Q.FindField('CTI_AXE3'         ).AsString);
        T.PutValue('CTI_AUCHOIX'      , Q.FindField('CTI_AUCHOIX'      ).AsString);
        T.PutValue('CTI_CIRCUITBAP'   , Q.FindField('CTI_CIRCUITBAP'   ).AsString);
        T.PutValue('CTI_COMPTE'       , Q.FindField('CTI_COMPTE'       ).AsString);
        T.PutValue('CTI_EXCLUSION'    , Q.FindField('CTI_EXCLUSION'    ).AsString);
        Inc(p);
      finally
        Ferme(Q);
      end;
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  p := 0;

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end
  else if HShowMessage('0;' + Ecran.Caption + ';Êtes vous sûr de vouloir dupliquer la sélection ?;Q;YN;N;N;', '', '') <> mrYes then
    Exit;

  {On passe par une tob pour le ca où l'on ferait un Select All et donc que l'on
   bouclerait que le Query}
  TobVisa := TOB.Create('_TYPE', nil, -1);
  TobCode := TOB.Create('_TYPE', nil, -1);
  try
    try
      {Chargement des codes éxistants, pour la génération des nouveaux codes}
      TobCode.LoadDetailFromSQL('SELECT CTI_CODEVISA FROM CPTYPEVISA');
      LastCode := '';

      {$IFNDEF EAGLCLIENT}
      TFMul(Ecran).Q.First;
      if F.AllSelected then
        while not TFMul(Ecran).Q.EOF do begin
          {Création d'un nouvel enregistrement}
          if not _DupliqueTypeVisa then Break;
          TFMul(Ecran).Q.Next;
        end
      else
      {$ENDIF}

      for n := 0 to F.nbSelected - 1 do begin
        F.GotoLeBookmark(n);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
        {$ENDIF}
        {Création d'un nouvel enregistrement}
        if not _DupliqueTypeVisa then Break;
      end;

      {Insertion en table}
      TobVisa.InsertDB(nil);

    except
      on E : Exception do begin
        RollBack;
        HShowMessage('0;' + Ecran.Caption + '; Traitement interrompu :'#13 + E.Message + ';E;O;O;O;', '', '');
      end;
    end;
  finally
    FreeAndNil(TobVisa);
    FreeAndNil(TobCode);
  end;

  if p > 0 then
    HShowMessage('0;' + Ecran.Caption + ';' + IntToStr(p) + ' type(s) de visas a (ont) été créé(s).;I;O;O;O;', '', '');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

initialization
  RegisterClasses([TOF_CPMULTYPEVISA]);

end.
