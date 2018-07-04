{***********UNITE*************************************************
Auteur  ...... : SG6
Créé le ...... : 08/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPGESTIONANA ()
Mots clefs ... : TOF;CPGESTIONANA
*****************************************************************}
unit CPGESTIONANA_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  dialogs,
  paramsoc, //GetParamSocSecur
  SAISUTIL, //TOBM
  HTB97,
  ed_tools,
  UtilSais,
  CPGENERAUX_Tom,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul,
  fe_main,
  {$ELSE}
  eMul,
  maineagl,
  //  main,
  {$ENDIF}
  LetBatch, //SG6 30/12/20014 VentileGenerale
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  utob,
  ent1,
  UTOF;

type
  TOF_CPGESTIONANA = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    //Contrôle de la fiche
    AXE1, AXE2, AXE3, AXE4, AXE5, TEMPCHECK, CROISAXE: TCheckBox;
    bValider: TToolBarButton97;

    //autre
    inChargement: boolean;

    //Evts sur les controles
    procedure CROISAXEClick(Sender: TObject);
    procedure bValiderClick(Sender: TObject);
    procedure AXEOnclick(Sender: TObject);

    //Function
    function RetournBoolForBase(Bool: boolean): string;

  end;


procedure CPLanceGestionAna();

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ULibanalytique; //SG6 22.03.05



{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 08/12/2004
Modifié le ... :   /  /
Description .. : Point d'entrée de la fiche CPGESTIONANA
Mots clefs ... : POINT ENTREE
*****************************************************************}
procedure CPLanceGestionAna();
begin
  AglLanceFiche('CP', 'CPGESTIONANA', '', '', '');
end;

procedure TOF_CPGESTIONANA.OnNew;
begin
  inherited;
end;

procedure TOF_CPGESTIONANA.OnDelete;
begin
  inherited;
end;

procedure TOF_CPGESTIONANA.OnUpdate;
begin
  inherited;
end;

procedure TOF_CPGESTIONANA.OnLoad;
begin
  inherited;
end;

procedure TOF_CPGESTIONANA.OnArgument(S: string);
begin
  inherited;
  inChargement := True;

  //Récupération des controles
  AXE1 := TCheckBox(getcontrol('AXE1', true));
  AXE2 := TCheckBox(getcontrol('AXE2', true));
  AXE3 := TCheckBox(getcontrol('AXE3', true));
  AXE4 := TCheckBox(getcontrol('AXE4', true));
  AXE5 := TCheckBox(getcontrol('AXE5', true));
  TEMPCHECK := TCheckBox(getcontrol('TEMPCHECK', true));
  CROISAXE := TCheckBox(getcontrol('CROISAXE', true));
  bValider := TToolBarButton97(getcontrol('bValider', true));


  //Gestion des evts sur les contrôles
  CROISAXE.OnClick := CROISAXEClick;
  AXE1.OnClick := AXEOnclick;
  AXE2.OnClick := AXEOnclick;
  AXE3.OnClick := AXEOnclick;
  AXE4.OnClick := AXEOnclick;
  AXE5.OnClick := AXEOnclick;
  bValider.OnClick := bValiderClick;

  CROISAXE.Checked := GetParamSocSecur('SO_CROISAXE',False);
  TEMPCHECK.Checked := GetParamSocSecur('SO_CROISAXE',False);
  if CROISAXE.Checked then
  begin
    AXE1.Checked := GetParamSocSecur('SO_VENTILA1',False);
    AXE2.Checked := GetParamSocSecur('SO_VENTILA2',False);
    AXE3.Checked := GetParamSocSecur('SO_VENTILA3',False);
    AXE4.Checked := GetParamSocSecur('SO_VENTILA4',False);
    AXE5.Checked := GetParamSocSecur('SO_VENTILA5',False);
  end
  else
  begin
    AXE1.Enabled := False;
    AXE2.Enabled := False;
    AXE3.Enabled := False;
    AXE4.Enabled := False;
    AXE5.Enabled := False;
  end;


  inChargement := False;
end;

procedure TOF_CPGESTIONANA.OnClose;
begin
  inherited;
end;

procedure TOF_CPGESTIONANA.OnDisplay();
begin
  inherited;
end;

procedure TOF_CPGESTIONANA.OnCancel();
begin
  inherited;
end;

procedure TOF_CPGESTIONANA.CROISAXEClick(Sender: TObject);
begin
  if inChargement then exit;

  AXE1.Enabled := CROISAXE.Checked;
  AXE2.Enabled := CROISAXE.Checked;
  AXE3.Enabled := CROISAXE.Checked;
  AXE4.Enabled := CROISAXE.Checked;
  AXE5.Enabled := CROISAXE.Checked;
  if not (CROISAXE.Checked) then
  begin
    AXE1.Checked := False;
    AXE2.Checked := False;
    AXE3.Checked := False;
    AXE4.Checked := False;
    AXE5.Checked := False;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 09/12/2004
Modifié le ... :   /  /
Description .. : Validatiion des paramètres
Suite ........ : - Suppression des écritures analytques
Suite ........ : - Mises à jour des paramsoc
Suite ........ : - Mise à jour des comptes généraux
Mots clefs ... : MAJ PARAMETRE GESTION ANALYTIQUE
*****************************************************************}
procedure TOF_CPGESTIONANA.bValiderClick(Sender: TObject);
var
  au_moins_unaxe, condition: boolean;
  stSet: string;
  lStWhereDelAnaPur : string;
  tobecr, vTobAna, tobecrglobale : TOB;
  traitement : TTraitementCroisaxe;
  lStrPieceAna  : string;
  lIntCpteur : longint;
  lSql : string;
  RecAxe : array[1..MaxAxe] of Boolean; {JP 02/08/07}
  n : Integer;
begin
  condition := ((GetParamSocSecur('SO_VENTILA1',False) = AXE1.Checked) and (GetParamSocSecur('SO_VENTILA2',False) = AXE2.Checked) and (GetParamSocSecur('SO_VENTILA3',False) = AXE3.Checked) and
    (GetParamSocSecur('SO_VENTILA4',False) = AXE4.Checked) and (GetParamSocSecur('SO_VENTILA5',False) = AXE5.Checked) and (GetParamSocSecur('SO_CROISAXE',False) = CROISAXE.Checked));
  if condition then
  begin
    Exit;
  end;
  au_moins_unaxe := AXE1.Checked or AXE2.Checked or AXE3.Checked or AXE4.Checked or AXE5.Checked;
  if CROISAXE.Checked then
  begin
    if not au_moins_unaxe then
    begin
      hshowmessage('0;Attention;Vous avez choisi le mode Croise-Axe mais aucun axe n''est séléctionné !;E;O;O;O;', '', '');
      exit;
    end;

    {FQ20900  20.120.2007  YMO}
    if (HShowMessage('0;' + Ecran.Caption + ';Attention, cette option ne peut être activée si vous utilisez le module gestion commerciale de CEGID BUSINESS PLACE. Voulez-vous continuer ?;Q;YN;N;N;', '', '') = mrNo)
    or (HShowMessage('0;' + Ecran.Caption + ';Êtes vous sûr de vouloir continuer ?;Q;YN;N;N;', '', '') = mrNo) then
      exit;
  end;
  if
    HShowMessage('0;Attention;Les ventilations analytiques affectées à des axes que vous ne choisissez pas pour le traitement en mode croise-axe seront définitivement perdues.#13#10Confirmez vous l''opération ?;Q;YN;N;N;', '', '')
    <>
    mrYes then exit;
  if HShowMessage('0;Attention;Vous pouvez faire une sauvegarde de la base pour pouvoir revenir en arrière.#13#10Voulez-vous interrompre le traitement ?;Q;YN;N;N;', '', '')
    <> mrNo then exit;
  //Traitement

  try
    begintrans;

    //Mise à jour des comptes généraux
    if CROISAXE.Checked then stSet := ''
                        else stSet := ' , G_VENTILABLE="-" ';

    if CROISAXE.Checked then
      ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE1="' + RetournBoolForBase(AXE1.Checked) + '" , G_VENTILABLE2="' + RetournBoolForBase(AXE2.Checked) + '" , G_VENTILABLE3="' +
        RetournBoolForBase(AXE3.Checked) + '" , G_VENTILABLE4="' + RetournBoolForBase(AXE4.Checked) + '" , G_VENTILABLE5="' + RetournBoolForBase(AXE5.Checked) + '"' + stSet +
        'WHERE G_VENTILABLE="X"');

    {JP 02/08/07 : Lorsque l'on décroise, on n'a besoin de savoir sur combien d'axes on décroise :
                   On stocke donc le paramétrage du CroiseAxe}
    for n := 1 to MaxAxe do
      RecAxe[n] := GetParamSocSecur('SO_VENTILA' + IntToStr(n), False);

    //Renseignements des paramsoc
    SetParamSoc('SO_CROISAXE', CROISAXE.Checked);
    SetParamSoc('SO_VENTILA1', AXE1.Checked);
    SetParamSoc('SO_VENTILA2', AXE2.Checked);
    SetParamSoc('SO_VENTILA3', AXE3.Checked);
    SetParamSoc('SO_VENTILA4', AXE4.Checked);
    SetParamSoc('SO_VENTILA5', AXE5.Checked);

    lStWhereDelAnaPur := '';
    if AXE1.Checked then
    begin
      lStWhereDelAnaPur := lStWhereDelAnaPur + ' AND Y_AXE<>"A1"';
    end;
    if AXE2.Checked then
    begin
      lStWhereDelAnaPur := lStWhereDelAnaPur + ' AND Y_AXE<>"A2"';
    end;
    if AXE3.Checked then
    begin
      lStWhereDelAnaPur := lStWhereDelAnaPur + ' AND Y_AXE<>"A3"';
    end;
    if AXE4.Checked then
    begin
      lStWhereDelAnaPur := lStWhereDelAnaPur + ' AND Y_AXE<>"A4"';
    end;
    if AXE5.Checked then
    begin
      lStWhereDelAnaPur := lStWhereDelAnaPur + ' AND Y_AXE<>"A5"';
    end;

    if CROISAXE.checked then
    begin
      //Suppression des ecritures od ana qui utilisé un axe non paramètré en mode croisaxe
      ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_TYPEANALYTIQUE="X"' + lStWhereDelAnaPur);

      //Maj ventilation
      tobecr := TOB.Create('ECRITURE', nil, -1);
      tobecrglobale := TOB.Create('$ECRITURE', nil, -1);
      traitement := TTraitementCroisaxe.Create(AXE1.Checked, AXE2.Checked, AXE3.Checked, AXE4.Checked, AXE5.Checked, False);

      {JP 07/11/05 : Croisement du paramétrage}
      Traitement.CroiseVentilTypeEtAutres;

      lSql := 'select ecriture.E_JOURNAL, ecriture.E_EXERCICE, ecriture.E_DATECOMPTABLE, ecriture.E_NUMEROPIECE, ecriture.E_NUMLIGNE, ecriture.E_QUALIFPIECE, ecriture.E_NUMECHE ';
      lSql := lSql + 'from generaux left join ecriture on G_GENERAL=E_GENERAL where G_VENTILABLE="X" AND E_GENERAL<>"" AND E_GENERAL IS NOT NULL';
      tobecrglobale.LoadDetailFromSQL(lSql);
      for lIntCpteur := 0 to tobecrglobale.Detail.Count - 1 do
      begin
        tobecr.SetString('E_JOURNAL', tobecrglobale.Detail[lIntCpteur].GetString('E_JOURNAL'));
        tobecr.SetString('E_EXERCICE', tobecrglobale.Detail[lIntCpteur].GetString('E_EXERCICE'));
        tobecr.SetDateTime('E_DATECOMPTABLE', tobecrglobale.Detail[lIntCpteur].GetDateTime('E_DATECOMPTABLE'));
        tobecr.SetInteger('E_NUMEROPIECE', tobecrglobale.Detail[lIntCpteur].GetInteger('E_NUMEROPIECE'));
        tobecr.SetInteger('E_NUMLIGNE', tobecrglobale.Detail[lIntCpteur].GetInteger('E_NUMLIGNE'));
        tobecr.SetInteger('E_NUMECHE', tobecrglobale.Detail[lIntCpteur].GetInteger('E_NUMECHE'));
        tobecr.SetString('E_QUALIFPIECE', tobecrglobale.Detail[lIntCpteur].GetString('E_QUALIFPIECE'));
        tobecr.LoadDB;
        CChargeAna(tobecr);
        if (tobecr.Detail.Count = 0) or ( {JP 17/12/07 : FQ 22060 : pour éviter un indice de liste hors limite sur tobecr.Detail[0]}
           (tobecr.Detail[0].Detail.Count = 0) and
           (tobecr.Detail[1].Detail.Count = 0) and
           (tobecr.Detail[2].Detail.Count = 0) and
           (tobecr.Detail[3].Detail.Count = 0) and
           (tobecr.Detail[4].Detail.Count = 0)) then
        begin
          tobecr.ClearDetail;
          continue;
        end;
        traitement.SetTobEcr(tobecr);
        traitement.PassageVentAnaClassiqueCroisaxe;
        traitement.DeleteInsertAna;
        tobecr.ClearDetail;
      end;
      tobecrglobale.Free;
      tobecr.Free;

      //Maj ana pur
      tobecrglobale := TOB.Create('$ANALYTIQ', nil, -1);
      vTobAna := TOB.Create('ANALYTIQ', nil, -1);
      tobecrglobale.LoadDetailFromSQL('SELECT Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE FROM ANALYTIQ WHERE Y_TYPEANALYTIQUE="X"');

      for lIntCpteur := 0 to tobecrglobale.Detail.Count - 1 do
      begin
        vTobAna.SetString('Y_JOURNAL', tobecrglobale.Detail[lIntCpteur].GetString('Y_JOURNAL'));
        vTobAna.SetString('Y_EXERCICE', tobecrglobale.Detail[lIntCpteur].GetString('Y_EXERCICE'));
        vTobAna.SetDateTime('Y_DATECOMPTABLE', tobecrglobale.Detail[lIntCpteur].GetDateTime('Y_DATECOMPTABLE'));
        vTobAna.SetInteger('Y_NUMEROPIECE', tobecrglobale.Detail[lIntCpteur].GetInteger('Y_NUMEROPIECE'));
        vTobAna.SetString('Y_AXE', tobecrglobale.Detail[lIntCpteur].GetString('Y_AXE'));
        vTobAna.SetString('Y_QUALIFPIECE', tobecrglobale.Detail[lIntCpteur].GetString('Y_QUALIFPIECE'));
        vTobAna.SetInteger('Y_NUMVENTIL', tobecrglobale.Detail[lIntcpteur].GetInteger('Y_NUMVENTIL'));
        vTobAna.SetInteger('Y_NUMLIGNE', tobecrglobale.Detail[lIntcpteur].GetInteger('Y_NUMLIGNE'));
        vTobAna.LoadDB;
        traitement.SetTobEcr(vTobAna);
        traitement.PassageAnaPurClassiqueCroisaxe;
        traitement.DeleteInsertAna;
      end;
      vTobAna.Free;
      traitement.MajSoldeSection;

      //Fin traitement analytique
      traitement.free;
    end
    else
    begin
      //Maj ventilation

      tobecr := TOB.Create('ECRITURE', nil, -1);
      tobecrglobale := TOB.Create('$ECRITURE', nil, -1);
      traitement := TTraitementCroisaxe.Create(False);

      {JP 02/08/07 : on mémorise les axes que l'on va traiter}
      Traitement.MajAxeTraitement(RecAxe);

      {JP 08/11/05 : Tant que le décroisement ne se fera pas, on supprimera les ventilations types}
      Traitement.SupprimeVentiType(True);

      lSql := 'select ecriture.E_JOURNAL, ecriture.E_EXERCICE, ecriture.E_DATECOMPTABLE, ecriture.E_NUMEROPIECE, ecriture.E_NUMLIGNE, ecriture.E_QUALIFPIECE, ecriture.E_NUMECHE ';
      lSql := lSql + 'from generaux left join ecriture on G_GENERAL=E_GENERAL where G_VENTILABLE="X" AND E_GENERAL<>"" AND E_GENERAL IS NOT NULL';
      tobecrglobale.LoadDetailFromSQL(lSql);

      for lIntCpteur := 0 to tobecrglobale.Detail.Count - 1 do
      begin
        tobecr.SetString('E_JOURNAL', tobecrglobale.Detail[lIntCpteur].GetString('E_JOURNAL'));
        tobecr.SetString('E_EXERCICE', tobecrglobale.Detail[lIntCpteur].GetString('E_EXERCICE'));
        tobecr.SetDateTime('E_DATECOMPTABLE', tobecrglobale.Detail[lIntCpteur].GetDateTime('E_DATECOMPTABLE'));
        tobecr.SetInteger('E_NUMEROPIECE', tobecrglobale.Detail[lIntCpteur].GetInteger('E_NUMEROPIECE'));
        tobecr.SetInteger('E_NUMLIGNE', tobecrglobale.Detail[lIntCpteur].GetInteger('E_NUMLIGNE'));
        tobecr.SetInteger('E_NUMECHE', tobecrglobale.Detail[lIntCpteur].GetInteger('E_NUMECHE'));
        tobecr.SetString('E_QUALIFPIECE', tobecrglobale.Detail[lIntCpteur].GetString('E_QUALIFPIECE'));
        tobecr.LoadDB;
        CChargeAna(tobecr);
        if (tobecr.Detail[0].Detail.Count = 0) and
           (tobecr.Detail[1].Detail.Count = 0) and
           (tobecr.Detail[2].Detail.Count = 0) and
           (tobecr.Detail[3].Detail.Count = 0) and
           (tobecr.Detail[4].Detail.Count = 0) then
        begin
          tobecr.ClearDetail;
          continue;
        end;
        traitement.SetTobEcr(tobecr);
        traitement.PassageVentCroisaxeAnaClassique;
        traitement.DeleteInsertAna;
        tobecr.ClearDetail;
      end;

      tobecrglobale.Free;
      tobecr.Free;

      //Maj ana pur
      tobecrglobale := TOB.Create('$ANALYTIQ', nil, -1);
      vTobAna := TOB.Create('$ANALYTIQ', nil, -1);

      lSql := 'SELECT Y_NUMEROPIECE, Y_EXERCICE, Y_DATECOMPTABLE, Y_QUALIFPIECE, Y_JOURNAL FROM ANALYTIQ WHERE Y_TYPEANALYTIQUE="X" ';
      lSql := lSql + 'AND Y_NUMVENTIL=1 ORDER BY Y_NUMEROPIECE, Y_NUMVENTIL';
      tobecrglobale.LoadDetailFromSQL(lSql);

      //On recupere les pieces d'od ana
      for lIntCpteur := 0 to tobecrglobale.Detail.Count - 1 do
      begin
        lStrPieceAna := 'SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="' + tobecrglobale.Detail[lIntCpteur].GetString('Y_JOURNAL') + '" ' +
         'AND Y_EXERCICE="' + tobecrglobale.Detail[lIntCpteur].GetString('Y_EXERCICE') + '" ' +
         'AND Y_DATECOMPTABLE="' + usDateTime(tobecrglobale.Detail[lIntCpteur].GetDateTime('Y_DATECOMPTABLE')) + '" ' +
         'AND Y_NUMEROPIECE=' + intToStr(tobecrglobale.Detail[lIntCpteur].GetInteger('Y_NUMEROPIECE')) + ' ' +
         'AND Y_NUMLIGNE=0 ' +
         'AND Y_QUALIFPIECE="' + tobecrglobale.Detail[lIntCpteur].GetString('Y_QUALIFPIECE') + '" ';

        vTobAna.LoadDetailDBFromSQL('ANALYTIQ', lStrPieceAna);
        traitement.SetTobEcr(vTobAna);
        traitement.PassageAnaPurCroisaxeClassique;
        traitement.DeleteInsertAna;
        vTobAna.ClearDetail;
       end;

      //Fin traitement analytique
      vTobAna.Free;
      tobecrglobale.Free;
      traitement.MajSoldeSection;
      traitement.free;
    end;

    CommitTrans;
    //SG6 26.01.05
    VH^.AnaCroisaxe := CROISAXE.Checked;

  except
    on e: exception do
    begin
      rollback;
      VH^.AnaCroisaxe := GetParamSocSecur('SO_CROISAXE',False);
      hshowmessage('E;Erreur;' + e.Message + ';E;O;O;O;', '', '');
      Exit;
    end;
  end;

  PGIInfo('Traitement effectué','Changement de mode analytique');

end;

{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 09/12/2004
Modifié le ... :   /  /
Description .. : Fonction qui retourne X si le paramètre est TRUE et
Suite ........ : retourne - sir le paramètre est FALSE
Mots clefs ... : BOOLEAN BDD
*****************************************************************}
function TOF_CPGESTIONANA.RetournBoolForBase(Bool: boolean): string;
begin
  if Bool then result := 'X' else result := '-';
end;

procedure TOF_CPGESTIONANA.AXEOnclick(Sender: TObject);
var
  num_axe: string;
begin
  if TCheckBox(Sender).Checked = False then Exit;
  num_axe := Copy(TCheckBox(Sender).Name, 4, 1);
  if not (ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_AXE="A' + num_axe + '"')) then
  begin
    hshowmessage('0;Attention;L''axe ' + num_axe + ' ne possède aucun code section et n''a pas de sections d''attente.;E;O;O;O;', '', '');
    TcheckBox(Sender).Checked := False;
  end;
end;


initialization
  registerclasses([TOF_CPGESTIONANA]);
end.

