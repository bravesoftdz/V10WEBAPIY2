{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/08/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PGRETROACTIV (PGRETROACTIV)
Mots clefs ... : TOM;PGRETROACTIV
*****************************************************************
PT1   24/08/2007 FC  V_72 : FQ 14682 ajouter des contrôles
PT2   28/08/2007 FC  V_72 : FQ 14689 ajouter un contrôle lors de la création d'une rétroactivité
}
unit UtomPGRETROACTIV;

interface

uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  Fiche,
  hdb,
  Fe_Main,
  FichList,
{$ELSE}
  MainEagl,
  eFiche,
  eFichList,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  Vierge,
  Windows,
  HTB97,
  pgoutils
  ;

type
  TOM_PGRETROACTIV = class(TOM)
  public

    bModeSaisie: Boolean;
    bModeCreation: Boolean;
    btRefresh: TToolbarButton97;
    NatureRub: integer;


    DateDebutExer: TDateTime;
    DateFinExer: TDateTime;

    Tob_cotisation: TOB;
//    Tob_EleNationaux    : TOB;
//    Tob_VariablePaie    : TOB;
    Tob_Remuneration: TOB;

    CEG, STD, DOS, OnFerme: Boolean; //PT1
    procedure RubriqueOnChange(Sender: Tobject);
    procedure ChercheInfo(Sender: Tobject);
    procedure RazChamps;
    procedure RendVisible(LaNature: Integer);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function RendValeurElt(TOBELT: TOB; EltNat: string; DD: TDateTime): string;
  public
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnAfterDeleteRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  end;

procedure PgRetroactiv_LanceFiche(Arguments: string; Index: string = '');

implementation

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 01/12/2006
Modifié le ... :   /  /
Description .. : lancement de la fiche de saisie des cotisations ou des
Suite ........ : rémunération.
Suite ........ : selon index, création ou modification
Mots clefs ... :
*****************************************************************}

procedure PgRetroactiv_LanceFiche(Arguments: string; Index: string = '');
begin
  AGLLanceFiche('PAY', 'PGRETROACTIV', '', Index, Arguments);
end;

procedure TOM_PGRETROACTIV.OnArgument(Arguments: string);
var
  Critere: string;
  mode: string;
  Q: Tquery;
  StrSql: string;
{$IFDEF EAGLCLIENT}
  Rubrique: THEdit;
{$ELSE}
  Rubrique: THDBEdit;
{$ENDIF EAGLCLIENT}
  Trubrique: THLabel;

begin
  inherited;
  bModeCreation := True;
  bModeSaisie := False;

  SetControlText('TLIBELLE', '');
  RazChamps;
  Critere := UpperCase_(Arguments);
  mode := READTOKENST(Critere);
  while mode <> '' do
  begin
    if mode = 'ACTION=MODIFICATION' then
    begin
      bModeCreation := False;
      bModeSaisie := True;
    end;

    if mode = 'COT' then
    begin
      NatureRub := 1;
{$IFDEF EAGLCLIENT}
      Rubrique := THEdit(GetControl('PGR_RUBRIQUE'));
{$ELSE}
      Rubrique := THDBEdit(GetControl('PGR_RUBRIQUE'));
{$ENDIF EAGLCLIENT}
      Rubrique.DataType := 'PGCOTIS';
      Rubrique.OnChange := RubriqueOnChange;
      Trubrique := THLabel(GetControl('TPGR_RUBRIQUE'));
      Trubrique.Caption := 'Cotisation';
      TFFiche(Ecran).caption := 'Saisie rétroactivité sur taux de cotisation : ';
      UpdateCaption(TFFiche(Ecran));
    end
    else
      if mode = 'AAA' then
      begin
        NatureRub := 2;
{$IFDEF EAGLCLIENT}
        Rubrique := THEdit(GetControl('PGR_RUBRIQUE'));
{$ELSE}
        Rubrique := THDBEdit(GetControl('PGR_RUBRIQUE'));
{$ENDIF EAGLCLIENT}
        Rubrique.DataType := 'PGREMUNERATION';
        Rubrique.OnChange := RubriqueOnChange;
        Trubrique := THLabel(GetControl('TPGR_RUBRIQUE'));
        Trubrique.Caption := 'Rémunération';
        Tob_Remuneration := Tob.Create('REMUNERATION', nil, -1);
        StrSql := 'SELECT PRM_RUBRIQUE, PRM_LIBELLE, PRM_CODECALCUL, PRM_COEFFREM, PRM_TYPECOEFF, PRM_TAUXREM, PRM_TYPETAUX, ' +
          'PRM_TYPEBASE, PRM_BASEREM, PRM_TYPEMONTANT, PRM_MONTANT ' +
          'FROM REMUNERATION WHERE ##PRM_PREDEFINI##';
        Tob_Remuneration.LoadDetailFromSQL(StrSql);
//      Tob_VariablePaie        := Tob.Create('VARIABLEPAIE', nil, -1);
//      StrSql                  := 'SELECT * FROM VARIABLEPAIE WHERE ##PVA_PREDEFINI##';
//      Tob_VariablePaie.LoadDetailFromSQL(StrSql);

{      Tob_EleNationaux        := Tob.Create('ELTNATIONAUX', nil, -1);
      StrSql                  := 'SELECT PEL_CODEELT,PEL_DATEVALIDITE,PEL_MONTANTEURO FROM ELTNATIONAUX ' +
                                'WHERER ##PEL_PREDEFINI## GROUP BY PEL_CODEELT,PEL_MONTANTEURO,PEL_DATEVALIDITE ' +
                                'ORDER BY PEL_DATEVALIDITE DESC';
      Tob_EleNationaux.LoadDetailFromSQL(StrSql);
}
        TFFiche(Ecran).caption := 'Saisie rétroactivité sur rémunération : ';
        UpdateCaption(TFFiche(Ecran));
      end;

    mode := READTOKENST(Critere);
  end;

  SetControlVisible('BINSERT', False);

  if bModeCreation then
  begin
    SetControlVisible('BREFRESH', True);
    btRefresh := TToolbarButton97(GetControl('BREFRESH'));
    btRefresh.OnClick := ChercheInfo;
  end
  else
  begin
    SetControlVisible('BREFRESH', False);
  end;

  Q := OpenSQL('SELECT MAX(PEX_DATEDEBUT),MAX(PEX_DATEFIN) FROM EXERSOCIAL', True);
  DateDebutExer := Q.Fields[0].AsDateTime;
  DateFinExer := Q.Fields[1].AsDateTime;
  Ferme(Q);

  TFVierge(Ecran).OnKeyDown := FormKeyDown;

  Tob_cotisation := Tob.Create('COTISATION', nil, -1);
  StrSql := 'SELECT PCT_RUBRIQUE, PCT_NATURERUB, PCT_LIBELLE, PCT_TYPETAUXSAL, PCT_TAUXSAL, PCT_TYPETAUXPAT, PCT_TAUXPAT' +
    ' FROM COTISATION WHERE PCT_NATURERUB="COT" AND ##PCT_PREDEFINI##';
  Tob_cotisation.LoadDetailFromSQL(StrSql);

  //PT1
  AccesPredefini('TOUS', CEG, STD, DOS);

{
  Q := OpenSQL('SELECT PEL_CODEELT,PEL_DATEVALIDITE,PEL_MONTANTEURO FROM ELTNATIONAUX '+
              'GROUP BY PEL_CODEELT,PEL_MONTANTEURO,PEL_DATEVALIDITE '+
              'ORDER BY PEL_DATEVALIDITE DESC', True);
  Tob_EleNationaux := Tob.Create('ELTNATIONAUX',nil,-1);
  Tob_EleNationaux.LoadDetailDB('ELTNATIONAUX','','',Q,False,false);
  Ferme(Q);
}

end;


procedure TOM_PGRETROACTIV.OnNewRecord;
begin
  inherited;
  SetField('PGR_PREDEFINI', 'DOS');
  if NatureRub = 1 then
    SetField('PGR_NATURERUB', 'COT')
  else
    if NatureRub = 2 then
      SetField('PGR_NATURERUB', 'AAA');

  Setfield('PGR_DATEVALIDITE', DateToStr(FINDEMOIS(Now)));
  SetField('PGR_DATEDEBUT', DateToStr(DateDebutExer));
  SetField('PGR_DATEFIN', DateToStr(DateFinExer));


end;

procedure TOM_PGRETROACTIV.OnDeleteRecord;
var
  StrSql: string;
  LeWhere: string;
  Reponse: integer;
  Pred : String;
begin
  inherited;
  //DEB PT1
  Pred := GetField('PGR_PREDEFINI');
  if (Pred = 'CEG') and (CEG = FALSE) then
  begin
    LastErrorMsg := 'Vous ne pouvez pas détruire une règle de prédéfini CEGID';
    LastError := -1;
    exit;
  end;
  //FIN PT1

  LeWhere := 'PGT_NATURERUB="' + GetField('PGR_NATURERUB') + '" AND PGT_RUBRIQUE="' +
    GetField('PGR_RUBRIQUE') + '" AND PGT_DATEVALIDITE="' + USDATETIME(GetField('PGR_DATEVALIDITE')) + '"';

  StrSql := 'SELECT 1 FROM PGHISTRETRO WHERE ' + Lewhere + ' AND PGT_DATEINTEG<>""';
  if ExisteSQL(StrSql) then
  begin
    LastError := -1;
    LastErrorMsg := 'Attention ! Certains éléments de l''historique sont déjà intégrés dans la paie.#10#13' +
      'Suppression impossible !';
  end
  else
  begin
    StrSql := 'SELECT 1 FROM PGHISTRETRO WHERE ' + Lewhere;
    if ExisteSQL(StrSql) then
    begin
      Reponse := PGIAsk('Attention ! Il existe un historique pour cette règle.#10#13' +
        'Voulez vous vraiment supprimer cette règle ?', 'Suppression');
      if Reponse = mrNo then
      begin
        LastError := -1;
      end
      else
      begin
        StrSql := 'DELETE FROM PGHISTRETRO WHERE ' + Lewhere;
        ExecuteSQL(StrSql);
      end;
    end;
  end;
end;

procedure TOM_PGRETROACTIV.OnUpdateRecord;
var
  bOk: boolean;
  StrErr: string;
  Pred : String;
  Q : TQuery;
begin
  inherited;
  bOk := True;
  //DEB PT1
  Pred := GetField('PGR_PREDEFINI');
  if (Pred = 'CEG') and (CEG = FALSE) then
  begin
    bOk := False;
    StrErr := 'Vous ne pouvez pas créer de règle prédéfini CEGID'; //PT1
    Pred := '';
    SetControlProperty('PGR_PREDEFINI', 'Value', Pred);
  end;
  //FIN PT1

  if (bOk) then
  begin
    //DEB PT2
    Q := OpenSQL('SELECT PGR_DATEVALIDITE FROM PGRETROACTIV WHERE PGR_PREDEFINI="' + GetField('PGR_PREDEFINI') + '"' +
      ' AND PGR_NODOSSIER="' + GetField('PGR_NODOSSIER') + '" AND PGR_NATURERUB = "' + GetField('PGR_NATURERUB') + '"' +
      ' AND PGR_RUBRIQUE ="' + GetField('PGR_RUBRIQUE') + '" AND PGR_DATEVALIDITE >= "' + USDATETIME(DEBUTDEMOIS(GetField('PGR_DATEVALIDITE'))) + '"' +
      ' AND PGR_DATEVALIDITE <= "' + USDATETIME(FINDEMOIS(GetField('PGR_DATEVALIDITE'))) + '" ORDER BY PGR_DATEVALIDITE DESC',True);
    if not Q.Eof then
    begin
      bOk := False;
      if NatureRub = 1 then
        StrErr := 'Une rétroactivité existe déjà pour la cotisation ' + GetField('PGR_RUBRIQUE') + #13' en date de validité ' + DateTimeToStr(Q.FindField('PGR_DATEVALIDITE').AsDateTime)
      else
        StrErr := 'Une rétroactivité existe déjà pour la rémunération ' + GetField('PGR_RUBRIQUE') + #13' en date de validité ' + DateTimeToStr(Q.FindField('PGR_DATEVALIDITE').AsDateTime);
    end;
    Ferme(Q);

    if (bOk) then
    begin
      Q := OpenSQL('SELECT PGR_DATEVALIDITE,PGR_DATEDEBUT,PGR_DATEFIN FROM PGRETROACTIV WHERE PGR_PREDEFINI="' + GetField('PGR_PREDEFINI') + '"' +
        ' AND PGR_NODOSSIER="' + GetField('PGR_NODOSSIER') + '" AND PGR_NATURERUB = "' + GetField('PGR_NATURERUB') + '"' +
        ' AND PGR_RUBRIQUE ="' + GetField('PGR_RUBRIQUE') + '"' +
        ' AND PGR_DATEDEBUT <="' + USDATETIME(GetField('PGR_DATEDEBUT')) + '"' +
        ' AND PGR_DATEFIN >="' + USDATETIME(GetField('PGR_DATEFIN')) + '"' +
        ' AND PGR_DATEVALIDITE <> "' + USDATETIME(GetField('PGR_DATEVALIDITE')) + '"' +
        ' ORDER BY PGR_DATEVALIDITE DESC',True);
      if not Q.Eof then
      begin
        bOk := False;
        if NatureRub = 1 then
          StrErr := 'Il existe déjà une rétroactivité pour la cotisation ' + GetField('PGR_RUBRIQUE') +
            ' en date de validité ' + DateTimeToStr(Q.FindField('PGR_DATEVALIDITE').AsDateTime) +
            #13' pour la période d''application du ' + DateTimeToStr(Q.FindField('PGR_DATEDEBUT').AsDateTime)+
            ' au ' + DateTimeToStr(Q.FindField('PGR_DATEFIN').AsDateTime)
        else
          StrErr := 'Il existe déjà une rétroactivité pour la rémunération ' + GetField('PGR_RUBRIQUE') +
            ' en date de validité ' + DateTimeToStr(Q.FindField('PGR_DATEVALIDITE').AsDateTime) +
            #13' pour la période d''application du ' + DateTimeToStr(Q.FindField('PGR_DATEDEBUT').AsDateTime)+
            ' au ' + DateTimeToStr(Q.FindField('PGR_DATEFIN').AsDateTime);
      end;
      Ferme(Q);
    end;
    //FIN PT2

    if (bOk) then
      if GetField('PGR_RUBRIQUE') = '' then
      begin
        Bok := False;
        if NatureRub = 1 then
          StrErr := 'Il faut choisir une cotisation.'
        else
          StrErr := 'Il faut une rubrique.';
      end;

    if (bOk) and (not bModeSaisie) then
    begin
      bOk := False;
      StrErr := 'Les éléments ne sont pas saisis.';
    end;

    //gestion des taux de cotisations
    if not bModeSaisie then
    begin
      bOk := False;
      StrErr := 'Vous devez saisir les nouvelles valeurs applicables. Pour cela, cliquez sur le bouton [Calculer les différentiels]'; //PT1
    end;
  end;
  
  if (not bOk) then
  begin
    LastError := -1;
    LastErrorMsg := StrErr;
  end;

end;

procedure TOM_PGRETROACTIV.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_PGRETROACTIV.OnAfterDeleteRecord;
begin
  inherited;
end;

procedure TOM_PGRETROACTIV.OnLoadRecord;
begin
  inherited;
  SetControlEnabled('PGR_NATURERUB', False);
  if not bModeCreation then
    RendVisible(NatureRub);

end;

procedure TOM_PGRETROACTIV.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_PGRETROACTIV.OnClose;
begin
  inherited;
  Tob_cotisation.Free;
//  Tob_EleNationaux.Free;
  Tob_Remuneration.Free;
//  Tob_VariablePaie.Free;
end;

procedure TOM_PGRETROACTIV.OnCancelRecord;
begin
  inherited;
end;



procedure TOM_PGRETROACTIV.ChercheInfo(Sender: Tobject);
var
  LeWhere: string;

begin
  SetFocusControl('PGR_PREDEFINI');

  //on fixe en début et fin de mois
  SetField('PGR_DATEDEBUT', DateTimeToStr(DEBUTDEMOIS(StrToDateTime(GetField('PGR_DATEDEBUT')))));
  SetField('PGR_DATEFIN', DateTimeToStr(FINDEMOIS(StrToDateTime(GetField('PGR_DATEFIN')))));

  LeWhere := 'PHB_RUBRIQUE="' + GetField('PGR_RUBRIQUE') + '" AND PHB_NATURERUB="' + GetField('PGR_NATURERUB') + '"' +
    ' AND PHB_DATEDEBUT>="' + USDATETIME(GetField('PGR_DATEDEBUT')) + '"' +
    ' AND PHB_DATEFIN<="' + USDATETIME(GetField('PGR_DATEFIN')) + '"';

  if not ExisteSQL('SELECT 1 FROM HISTOBULLETIN WHERE ' + LeWhere) then
  begin
    PGIInfo('Il n''y a pas de mouvement pour cette rubrique à cette période.');
    SetFocusControl('PGR_DATEDEBUT');
    RazChamps;
    exit;
  end;

  SetControlText('TLIBELLE', '');
  SetControlText('TXRUBSAL', '0');
  SetControlText('TXRUBPAT', '0');
  SetField('PGR_DIFFTXSAL', 0);
  SetField('PGR_DIFFTXPAT', 0);
  SetField('PGR_DIFFBASECOT', 0);
  SetField('PGR_DIFFBASEREM', 0);
  SetField('PGR_DIFTAUXREM', 0);
  SetField('PGR_DIFCOEFFREM', 0);
  SetField('PGR_DIFFMTREM', 0);
  SetControlText('BASEREM', '0');
  SetControlText('TAUXREM', '0');
  SetControlText('COEFFREM', '0');
  SetControlText('MTREM', '0');
  RazChamps;
  bModeSaisie := True;

  RendVisible(NatureRub);

end;


procedure TOM_PGRETROACTIV.RazChamps;
begin
  //Cotisations
  SetControlVisible('TTXRUBSAL', False);
  SetControlVisible('TTXRUBPAT', False);
  SetControlVisible('TXRUBSAL', False);
  SetControlVisible('TXRUBPAT', False);
  SetControlVisible('PGR_DIFFTXSAL', False);
  SetControlVisible('TPGR_DIFFTXSAL', False);
  SetControlVisible('PGR_DIFFTXPAT', False);
  SetControlVisible('TPGR_DIFFTXPAT', False);

  //Rémunérations
  SetControlVisible('PGR_DIFFBASEREM', False);
  SetControlVisible('TPGR_DIFFBASEREM', False);
  SetControlVisible('PGR_DIFTAUXREM', False);
  SetControlVisible('TPGR_DIFTAUXREM', False);
  SetControlVisible('PGR_DIFCOEFFREM', False);
  SetControlVisible('TPGR_DIFCOEFFREM', False);
  SetControlVisible('PGR_DIFMTREM', False);
  SetControlVisible('TPGR_DIFMTREM', False);
  SetControlVisible('CODECALCUL', False);
  SetControlVisible('TCODECALCUL', False);
  SetControlVisible('BASEREM', False);
  SetControlVisible('TAUXREM', False);
  SetControlVisible('COEFFREM', False);
  SetControlVisible('MTREM', False);
  SetControlVisible('TPGR_DIFFREMANC', False);
  SetControlVisible('TPGR_DIFFREM', False);
  bModeSaisie := False;

end;


procedure TOM_PGRETROACTIV.RendVisible(LaNature: integer);
var
  TobCot: Tob;
  TobRem: Tob;
  TobNat: Tob;
  PasEln: integer;
  Q: TQuery;
  StrSql: string;
  Strin: array[0..4] of string;
  Index: integer;
  i: integer;
  DD: TDateTime;
  dt: TDateTime;
begin
  TobNat := Tob.Create('ELTNATIONAUX', nil, -1);
  if LaNature = 1 then //Cotisation
  begin
    PasEln := 0;
    TobCot := Tob_cotisation.FindFirst(['PCT_NATURERUB', 'PCT_RUBRIQUE'], [GetField('PGR_NATURERUB'), GetField('PGR_RUBRIQUE')], False);
    if Tobcot <> nil then
    begin
      if TobCot.GetValue('PCT_TYPETAUXSAL') = 'ELN' then
      begin
      dt := GetField('PGR_DATEVALIDITE');   //PT1
      //FormatDateTime('d mmmm yyyy',
        Q := OpenSQL('SELECT PEL_CODEELT,PEL_DATEVALIDITE,PEL_MONTANTEURO FROM ELTNATIONAUX ' +
          ' WHERE  ##PEL_PREDEFINI## AND PEL_CODEELT="' + TobCot.GetValue('PCT_TAUXSAL') + '" ' +
          ' AND PEL_DATEVALIDITE <= "' + USDATETIME(dt) + '" ' + //PT1
          ' GROUP BY PEL_CODEELT,PEL_MONTANTEURO,PEL_DATEVALIDITE ' +
          ' ORDER BY PEL_DATEVALIDITE DESC', True);
//        if TobCot.GetValue('PCT_TAUXSAL')<>'' then
        if not Q.Eof then
        begin
//          TobCalc:=Tob_EleNationaux.FindFirst(['PEL_CODEELT'],[TobCot.GetValue('PCT_TAUXSAL')],False);
//          SetControlText('TXRUBSAL', TobCalc.GetValue('PEL_MONTANTEURO'));
          SetControlText('TXRUBSAL', Q.FindField('PEL_MONTANTEURO').AsVariant);
          Ferme(Q);
        end;
      end
      else
        inc(PasEln);

      if TobCot.GetValue('PCT_TYPETAUXPAT') = 'ELN' then
      begin
        dt := GetField('PGR_DATEVALIDITE');   //PT1
        Q := OpenSQL('SELECT PEL_CODEELT,PEL_DATEVALIDITE,PEL_MONTANTEURO FROM ELTNATIONAUX ' +
          'WHERE ##PEL_PREDEFINI## AND PEL_CODEELT="' + TobCot.GetValue('PCT_TAUXPAT') + '" ' +
          ' AND PEL_DATEVALIDITE <= "' + USDATETIME(dt) + '" ' + //PT1
          'GROUP BY PEL_CODEELT,PEL_MONTANTEURO,PEL_DATEVALIDITE ' +
          'ORDER BY PEL_DATEVALIDITE DESC', True);
//        if TobCot.GetValue('PCT_TAUXPAT')<>'' then
        if not Q.Eof then
        begin
//          TobCalc:=Tob_EleNationaux.FindFirst(['PEL_CODEELT'],[TobCot.GetValue('PCT_TAUXPAT')],False);
//          SetControlText('TXRUBPAT', TobCalc.GetValue('PEL_MONTANTEURO'));
          SetControlText('TXRUBPAT', Q.FindField('PEL_MONTANTEURO').AsVariant);
          Ferme(Q);
        end;
      end
      else
        inc(PasEln);

      if PasEln = 2 then
      begin
        PGIInfo('Cette cotisation n''est pas élément national.');
        RazChamps;
        exit;
      end;


    end;
    SetControlVisible('TTXRUBSAL', True);
    SetControlVisible('TTXRUBPAT', True);
    SetControlVisible('TXRUBSAL', True);
    SetControlVisible('TXRUBPAT', True);
    SetControlVisible('PGR_DIFFTXSAL', True);
    SetControlVisible('TPGR_DIFFTXSAL', True);
    SetControlVisible('PGR_DIFFTXPAT', True);
    SetControlVisible('TPGR_DIFFTXPAT', True);

  {
    if not bModeCreation then
    begin
      NvTaux := GetField('PGR_DIFFTXSAL')+(StrToFloat(GetControlText('TXRUBSAL')));
      SetControlText('TXSALSAISIE', FloatToStr(NvTaux));
      NvTaux := GetField('PGR_DIFFTXPAT')+(StrToFloat(GetControlText('TXRUBPAT')));
      SetControlText('TXPATSAISIE', FloatToStr(NvTaux));
    end;
 }
    SetFocusControl('PGR_DIFFTXSAL');
  end
  else
    if NatureRub = 2 then //Rémunérations
    begin
      SetControlVisible('PGR_DIFFBASEREM', True);
      SetControlVisible('TPGR_DIFFBASEREM', True);
      SetControlVisible('PGR_DIFTAUXREM', True);
      SetControlVisible('TPGR_DIFTAUXREM', True);
      SetControlVisible('PGR_DIFCOEFFREM', True);
      SetControlVisible('TPGR_DIFCOEFFREM', True);
      SetControlVisible('PGR_DIFFMTREM', True);
      SetControlVisible('TPGR_DIFFMTREM', True);
      SetControlVisible('CODECALCUL', True);
      SetControlVisible('TCODECALCUL', True);
      SetControlEnabled('CODECALCUL', False);
      SetControlVisible('BASEREM', True);
      SetControlVisible('TAUXREM', True);
      SetControlVisible('COEFFREM', True);
      SetControlVisible('MTREM', True);
      SetControlVisible('TPGR_DIFFREMANC', True);
      SetControlVisible('TPGR_DIFFREM', True);

      TobRem := Tob_Remuneration.FindFirst(['PRM_RUBRIQUE'], [GetField('PGR_RUBRIQUE')], False);
      if TobRem <> nil then
      begin
        SetControlText('CODECALCUL', TobRem.getvalue('PRM_CODECALCUL'));
        SetField('PGR_DIFCOEFFREM', 0);
        SetField('PGR_DIFTAUXREM', 0);
        SetField('PGR_DIFFBASEREM', 0);
        SetField('PGR_DIFFMTREM', 0);
        SetControlText('BASEREM', '0');
        SetControlText('TAUXREM', '0');
        SetControlText('COEFFREM', '0');
        SetControlText('MTREM', '0');

        if TobRem.GetValue('PRM_CODECALCUL') = '01' then //montant
        begin
          SetControlEnabled('PGR_DIFFBASEREM', True);
          SetControlEnabled('PGR_DIFTAUXREM', False);
          SetControlEnabled('PGR_DIFCOEFFREM', False);
          SetControlEnabled('PGR_DIFFMTREM', True);
          SetControlEnabled('BASEREM', True);
          SetControlEnabled('TAUXREM', False);
          SetControlEnabled('COEFFREM', False);
          SetControlEnabled('MTREM', True);
        end;
        if (TobRem.GetValue('PRM_CODECALCUL') = '02') or (TobRem.GetValue('PRM_CODECALCUL') = '03')
          or (TobRem.GetValue('PRM_CODECALCUL') = '06') or (TobRem.GetValue('PRM_CODECALCUL') = '07') then //Base taux coef
        begin
          SetControlEnabled('PGR_DIFFBASEREM', True);
          SetControlEnabled('PGR_DIFTAUXREM', True);
          SetControlEnabled('PGR_DIFCOEFFREM', True);
          SetControlEnabled('PGR_DIFFMTREM', False);
          SetControlEnabled('BASEREM', True);
          SetControlEnabled('TAUXREM', True);
          SetControlEnabled('COEFFREM', True);
          SetControlEnabled('MTREM', False);
        end;
        if (TobRem.GetValue('PRM_CODECALCUL') = '04') or (TobRem.GetValue('PRM_CODECALCUL') = '05') then //Base taux
        begin
          SetControlEnabled('PGR_DIFFBASEREM', True);
          SetControlEnabled('PGR_DIFTAUXREM', True);
          SetControlEnabled('PGR_DIFCOEFFREM', False);
          SetControlEnabled('PGR_DIFFMTREM', False);
          SetControlEnabled('BASEREM', True);
          SetControlEnabled('TAUXREM', True);
          SetControlEnabled('COEFFREM', False);
          SetControlEnabled('MTREM', False);
        end;
        if TobRem.GetValue('PRM_CODECALCUL') = '08' then //coef
        begin
          SetControlEnabled('PGR_DIFFBASEREM', True);
          SetControlEnabled('PGR_DIFTAUXREM', False);
          SetControlEnabled('PGR_DIFCOEFFREM', True);
          SetControlEnabled('PGR_DIFFMTREM', False);
          SetControlEnabled('BASEREM', True);
          SetControlEnabled('TAUXREM', False);
          SetControlEnabled('COEFFREM', True);
          SetControlEnabled('MTREM', False);
        end;

        StrSql := 'SELECT PEL_CODEELT,PEL_DATEVALIDITE,PEL_MONTANTEURO FROM ELTNATIONAUX WHERE ##PEL_PREDEFINI## ';
        Index := 0;
        if TobRem.GetValue('PRM_TYPEBASE') = '02' then
        begin
          StrIn[Index] := '"' + TobRem.GetValue('PRM_BASEREM') + '"';
          Inc(Index);
        end;
        if TobRem.GetValue('PRM_TYPETAUX') = '02' then
        begin
          StrIn[Index] := '"' + TobRem.GetValue('PRM_TAUXREM') + '"';
          Inc(Index);
        end;
        if TobRem.GetValue('PRM_TYPECOEFF') = '02' then
        begin
          StrIn[Index] := '"' + TobRem.GetValue('PRM_COEFFREM') + '"';
          Inc(Index);
        end;
        if TobRem.GetValue('PRM_TYPEMONTANT') = '02' then
        begin
          StrIn[Index] := '"' + TobRem.GetValue('PRM_MONTANT') + '"';
          Inc(Index);
        end;
        if Index > 0 then
        begin
          StrSql := StrSql + ' AND (PEL_CODEELT IN (';
          for i := 0 to Index - 1 do
          begin
            StrSql := StrSql + StrIn[i];
            if i + 1 < Index then
              StrSql := StrSql + ', ';
          end;
          dt := GetField('PGR_DATEVALIDITE');  //PT1
          StrSql := StrSql + ')) AND PEL_DATEVALIDITE <= "' + USDATETIME(dt) + '" ' + //PT1
            ' GROUP BY PEL_CODEELT,PEL_MONTANTEURO,PEL_DATEVALIDITE ORDER BY PEL_DATEVALIDITE DESC';
          TobNat.LoadDetailFromSQL(StrSql);
        end;

        //paramétrage selon type 08 -> valeur, 03 -> variable paie
        if TobRem.GetValue('PRM_TYPEBASE') = '08' then
          SetControlText('BASEREM',  TobRem.GetValue('PRM_BASEREM'));

        if TobRem.GetValue('PRM_TYPECOEFF') = '08' then
          SetControlText('COEFFREM', TobRem.GetValue('PRM_COEFFREM'));

        if TobRem.GetValue('PRM_TYPETAUX') = '08' then
          SetControlText('TAUXREM', TobRem.GetValue('PRM_TAUXREM'));

        if TobRem.GetValue('PRM_TYPEMONTANT') = '08' then
          SetControlText('MTREM', TobRem.GetValue('PRM_MONTANT'));

        DD := GetField('PGR_DATEDEBUT');
        // paramètrage type 03 -> éléments nationaux
        if TobRem.GetValue('PRM_TYPEBASE') = '02' then
        begin
          SetControlText('BASEREM', RendValeurElt(TobNat, Tobrem.GetValue('PRM_BASEREM'), DD));
        end;

        if TobRem.GetValue('PRM_TYPETAUX') = '02' then
        begin
          SetControlText('TAUXREM', RendValeurElt(TobNat, Tobrem.GetValue('PRM_TAUXREM'), DD));
        end;

        if TobRem.GetValue('PRM_TYPECOEFF') = '02' then
        begin
          SetControlText('COEFFREM', RendValeurElt(TobNat, Tobrem.GetValue('PRM_COEFFREM'), DD));
        end;

        if TobRem.GetValue('PRM_TYPEMONTANT') = '02' then
        begin
          SetControlText('MTREM', RendValeurElt(TobNat, Tobrem.GetValue('PRM_MONTANT'), DD));
        end;
      end;
      SetFocusControl('PGR_DIFFBASEREM');
    end;
  TobNat.Free;
end;

procedure TOM_PGRETROACTIV.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F10: ChercheInfo(sender);
  end;
end;


procedure TOM_PGRETROACTIV.RubriqueOnChange(Sender: Tobject);
var
  Q: Tquery;

begin
  if GetControlText('PGR_NATURERUB') = 'COT' then
  begin
    Q := OpenSQL('SELECT PCT_LIBELLE FROM COTISATION WHERE PCT_RUBRIQUE="' + GetControlText('PGR_RUBRIQUE') +
      '" AND ##PCT_PREDEFINI##', True);
    SetControlText('TLIBELLE', Q.Fields[0].AsString);
    Ferme(Q);
  end;

  if GetControlText('PGR_NATURERUB') = 'AAA' then
  begin
    Q := OpenSQL('SELECT PRM_LIBELLE FROM REMUNERATION WHERE PRM_RUBRIQUE="' + GetControlText('PGR_RUBRIQUE') +
      '" AND ##PRM_PREDEFINI##', True);
    SetControlText('TLIBELLE', Q.Fields[0].AsString);
    Ferme(Q);
  end;

end;

function TOM_PGRETROACTIV.RendValeurElt(TOBELT: TOB; EltNat: string; DD: TDateTime): String;
var
  Bok: Boolean;
  TobCalc: TOB;
  ValDD: Double;
begin
  Result := '';
  ValDD := 0;
  // Boucle pour valorisation élément national par rapport à la date de début de la période à traiter
  Bok := False;
  TobCalc := TOBELT.FindFirst(['PEL_CODEELT'], [EltNat], False);
  if TobCalc <> nil then ValDD := TobCalc.GetValue('PEL_MONTANTEURO');
  while (not Bok) and (TobCalc <> nil) do
  begin
    if DD < TobCalc.GetValue('PEL_DATEVALIDITE') then
      Bok := true
    else
      TobCalc := TOBELT.FindNext(['PEL_CODEELT'], [EltNat], true);
  end;
  if TobCalc <> nil then ValDD := TobCalc.GetValue('PEL_MONTANTEURO');
  Result := FloatToStr(ValDD);
end;

initialization
  registerclasses([TOM_PGRETROACTIV]);
end.

