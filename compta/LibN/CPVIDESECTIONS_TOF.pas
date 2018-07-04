{ Unité : Source TOF de la FICHE : CPVIDESECTIONS
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 6.50.001.004    09/06/05    JP     Création de l'unité : migration eAgl de VideSect.Pas
--------------------------------------------------------------------------------------}
unit CPVIDESECTIONS_TOF;
{Principes du traitement :
 1/ Le paramétrage :
    a/ Clefs de répartition : On détermine pour un axe donné et éventuellement pour une sélection de
                              comptes généraux sur quelles sections on va faire la répartition et selon
                              quel mode de calcul : Quantité, section/période ... (cf CPREPARCLE_TOF)
    b/ Sur les sections : on peut rattacher une clef de répartition paramétrée ci-dessus

 2/ Le traitement :
    a/ Les critères :
       - On choisit la période et les comptes dont on va traiter les écitures analytiques
       - On détermine l'axe qui va servir de référence et la clef de répartition à appliquer
       - On choisit la période et les comptes sur lesquels on va faire le calcul du Chiffre d'Affaire
         sur les différentes sections de la clef de répartition
    b/ Paramètres : Ces zones permettent de définir les écritures des pièces d'OD analytiques que l'on va
                    générer : Y_JOURNAL, Y_LIBELLE, Y_REFINTERNE, Y_DATECOMPTABLE
    c/ Traitement : 1/ Dans un premier temps, on calcule le montant net des écritures analytiques sur la période
                       sélectionnée dont les sections on comme clef de répartition la clef sélectionnée.
                    2/ Ensuite on calcule le CA sur les sections de la clef de répartition correspondant aux
                       comptes et à la période de référence.
                    3/ Enfin, on génère des OD analytiques qui soldent les écitures de 1/ et on répartit les
                       montants au prorata des totaux par section trouvés en 2/

 3/ Exemple :
    On veut répartir les montants sur la section Attente des comptes 6 sur le mois de 06/2005
    La section Attente en question a une clef de répartition concernant les sections A, B et C
    On veut que la répartition se fasse sur le CA des comptes 7 sur le premier semestre 2005.
    On choisit donc la clef idoine afin de traiter les écritures de classe 6 dont la section Attente.
    Sur le mois de juin 2005 j'ai un montant net de 10000€ en crédit sur écritures analytiques de classe 6 avec la Section Attente
    Sur le premier semestre 2005, j'ai 5000€ de CA sur les comptes de classe 7 avec la section A (25%)
                                          0€ de CA sur les comptes de classe 7 avec la section B  (0%)
                                      15000€ de CA sur les comptes de classe 7 avec la section C (75%)
    Je vais donc générer une OD sur le journal choisi et à la date de génération
    Y_Debit     Y_Credit  Y_Section
     10000€                Attente
                 2500€          A
                 7500€          B
}

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db,  dbtables, FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  Forms, SysUtils, ComCtrls, HCtrls, uTob, UTOF, HTB97, Ent1,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  {Pour la gestion des filtres}
  UObjFiltres;

type
  P_RAP = class
    Section  : string;
    General  : string;
    Sousplan : array[1..MaxAxe] of string;
    Solde    : double;
  end;

  TOF_CPVIDESECTIONS = class (TOF)
    procedure OnUpdate             ; override;
    procedure OnLoad               ; override;
    procedure OnArgument(S :string); override;
    procedure OnClose              ; override;
  private
    ObjFiltre  : TObjFiltre;

    procedure GestionFlash(Fin : Boolean = True);
    procedure ViderListe(Detruire : Boolean = False);
    function  ControlDesZones : Boolean;
    function  ControleDate    : Boolean;
    function  ControleJournal : Boolean;
  protected
    DateGene   : TDateTime;
    NumAxe     : Integer;
    AxeJal     : string;
    WhereEmet  : string;
    WhereRecept: string;
    TSource    : TList;
    TDest      : TList;
    TPieces    : TList;
    vTobAna    : TOB;
    {Validation de la saisie avant le lancement du traitement}
    function ControleOK : Boolean;
    {Teste s'il y a déjà eu un transfert intersection pour ces critères}
    function DejaTIS    : Boolean;
    function CoherQte     (ModeR, QQ : string) : Boolean;
    function ChargeDest   (ModeR, QQ : string) : string;
    function FabricSQLAnal(ModeR, QQ, Sect : string; Emet : Boolean): string;
    function CreerLigneAna(NewNum, NumV : Longint; Section, General : string; Debit, Credit : Double; SousPlan : array of string) : TOB;

    procedure MajDernTIS;
    procedure GenereLesPieces;
    procedure ConstituteClausesWhere;
    procedure RemplirLaListe(Q : TQuery; ModeR : string; Source : Boolean);
  public
    procedure S_CLEREPARTITIONChange (Sender : TObject);
    procedure BGenereClick           (Sender : TObject);
    procedure Y_AXEChange            (Sender : TObject);
    procedure DATEGENERATIONExit     (Sender : TObject);
    procedure DATECOMPTABLEExit      (Sender : TObject);
    procedure Y_JOURNALChange        (Sender : TObject);
  end;

procedure VidageInterSections;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  SaisUtil, UtilPgi, DelVisuE, ed_tools, LettUtil, HEnt1, HMsgBox, HStatus, ULibAnalytique,
  UtilSais;

{---------------------------------------------------------------------------------------}
procedure VidageInterSections;
{---------------------------------------------------------------------------------------}
begin
  {On vérifie la compatibilité de la date d'entrée avec les exercices ouverts}
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  {On s'assure qu'il n'y a pas de blocage sur la base pour des traitements de masse}
  if _Blocage(['nrCloture', 'nrBatch', 'nrSaisieModif'], False, 'nrBatch') then Exit;
  {Appel de la fiche}
  try
    AglLanceFiche('CP', 'CPVIDESECTIONS', '', '', '');
  finally
    {Quoiqu'il arrive, on libère la base}
    _Bloqueur('nrBatch', False);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  Composant : TControlFiltre ;
begin
  inherited;
  Ecran.HelpContext := 7379000;
  {Gestion des filtres, centralisée dans UObjFiltres}
  Composant.Filtre   := TToolbarButton97(GetControl('BFILTRE'  ));
  Composant.Filtres  := THValComboBox   (GetControl('FFILTRES' ));
  Composant.PageCtrl := TPageControl    (GetControl('PCCONTROL'));
  ObjFiltre := TObjFiltre.Create(Composant, 'VIDESECT');

  {Branchement des évènements}
  THValComboBox(GetControl('Y_JOURNAL'       )).OnChange := Y_JOURNALChange;
  THValComboBox(GetControl('Y_AXE'           )).OnChange := Y_AXEChange;
  THValComboBox(GetControl('S_CLEREPARTITION')).OnChange := S_CLEREPARTITIONChange;
  THEdit       (GetControl('DATEGENERATION'  )).OnExit   := DATEGENERATIONExit;
  THEdit       (GetControl('Y_DATECOMPTABLE_')).OnExit   := DATECOMPTABLEExit;
  THEdit       (GetControl('Y_DATECOMPTABLE' )).OnExit   := DATECOMPTABLEExit;
  TToolbarButton97(GetControl('BGENERE'      )).OnClick  := BGenereClick;

  TSource := TList.Create;
  TDest   := TList.Create;
  TPieces := TList.Create;
  {Tob dans laquelle vont être stockées les OD analytiques issues du traitement}
  vTobAna := TOB.Create('$$ANALYTIQ', nil, -1);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.OnClose;
{---------------------------------------------------------------------------------------}
begin
  ViderListe(True);
  if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetControlText('Y_AXE', 'A1');
  Y_AXEChange(GetControl('Y_AXE'));

  DateGene := V_PGI.DateEntree;
  AxeJal := '';
  ObjFiltre.Charger;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  St, QQ: string;
  Q, QY: TQuery;
  ModeR: string;
  Tot: double;
  i: integer;
begin
  inherited;
  {On s'assure que les zones ont bien été renseignées}
  if not ControleOK then Exit;
  {Demande de confirmation du traitement}
  if HShowMessage('11;' + Ecran.Caption + ';Confirmez-vous la génération du transfert et des pièces associées ?;Q;YN;N;N;', '', '') <> mrYes then Exit;
  {Si un transfert a déjà exécuté}
  if DejaTIS then Exit;
  {Génération des clauses where sur les généraux}
  ConstituteClausesWhere;

  GestionFlash(False);
  try
    {Lecture des infos de la clef de répartition}
    Q := OpenSQL('SELECT RE_MODEREPARTITION, RE_QUALIFQTE FROM CLEREPAR WHERE ' +
                 'RE_AXE = "' + GetControlText('Y_AXE') + '" AND RE_CLE = "' + GetControlText('S_CLEREPARTITION') + '"', True);
    if not Q.EOF then begin
      ModeR := Q.FindField('RE_MODEREPARTITION').AsString;
      QQ    := Q.FindField('RE_QUALIFQTE').AsString;
      Ferme(Q);
      {Vérification des quantités}
      if not CoherQte(ModeR, QQ) then begin
        if HShowMessage('13;' + Ecran.Caption + ';Certains qualifiants quantités des mouvements sont incohérents. '#13 +
                        'Voulez-vous lancer le traitement sur les qualifiants correspondant à la clé de répartition ?;Q;YN;Y;Y;', '', '') <> mrYes then
          Exit;
      end;
      {Construction de la requête}
      St := FabricSQLAnal(ModeR, QQ, '', True);
    end
    else begin
      Ferme(Q);
      Exit;
    end;

    {On vide les listes et vTobAna avant le lancement du traitement proprement dit}
    ViderListe;

    {La requête n'est peut être pas correcte vu le nombre de concaténations dont elle est issues}
    try
      QY := OpenSQL(St, True);
    except
      on E : Exception do begin
        MessageAlerte(E.Message + #13#13 + 'Vérifiez la définition des critères sur les comptes généraux.');
        Exit;
      end;
    end;

    {Chargement des infos emettrices}
    InitMove(QY.RecordCount, 'Chargement des infos émettrices');
    try
      while not QY.EOF do begin
        {On remplit la liste TSource}
        MoveCur(False);
        RemplirLaListe(QY, ModeR, True);
        QY.Next;
      end;
    finally
      Ferme(QY);
      FiniMove;
    end;

    if TSource.Count <= 0 then begin
      GestionFlash;
      HShowMessage('8;' + Ecran.Caption + ';Il n''existe aucun mouvement analytique pour ces critères;E;O;O;O;', '', '');
      Exit;
    end;

    {Lecture de la requete et chargement des infos receptrices}
    ChargeDest(ModeR, QQ);
    if TDest.Count <= 0 then begin
      HShowMessage('9;' + Ecran.Caption + ';Il n''existe aucune section de transfert pour cette clé de répartition;E;O;O;O;', '', '');
      Exit;
    end;

    {Calcul de la somme des mouvements}
    Tot := 0;
    for i := 0 to TDest.Count - 1 do
      Tot := Tot + P_RAP(TDest[i]).Solde;

    if Tot = 0 then begin
      HShowMessage('10;' + Ecran.Caption + ';La somme des mouvements sur les sections receptrices est nulle;E;O;O;O;', '', '');
      Exit;
    end;

    {Génération des pièces}
    if Transactions(GenereLesPieces, 0) <> oeOK then
      HShowMessage('14;' + Ecran.Caption + ';ATTENTION! Génération non effectuée.;E;O;O;O;', '', '')
    else begin
      if Transactions(MajDernTIS, 0) <> oeOk then
      HShowMessage('16;' + Ecran.Caption + ';ATTENTION! Génération non effectuée mais non mémorisée.;E;O;O;O;', '', '');
      if TPieces.Count > 0 then begin
        if HShowMessage('12;' + Ecran.Caption + ';Voulez-vous voir la liste des pièces générées?;Q;YN;Y;Y;', '', '') = mrYes then
          VisuPiecesGenere(TPieces, EcrAna, 5);
      end;
    end;
  finally
    GestionFlash;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.BGenereClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if TPieces.Count > 0 then
    VisuPiecesGenere(TPieces, EcrAna, 5);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.DATEGENERATIONExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  ControleDate;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.DATECOMPTABLEExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Par défaut, la date de génération est la date de la fin de période de référence}
  if IsValidDate(GetControlText('Y_DATECOMPTABLE_')) then
    SetControlText('DATEGENERATION', GetControlText('Y_DATECOMPTABLE_'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.S_CLEREPARTITIONChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if GetControlText('S_CLEREPARTITION') = '' then begin
    SetControlText('GENERECEPT', '');
    Exit;
  end;

  {Récupération des comptes généraux paramétrés dans la répartition des clefs analytiques}
  Q := OpenSQL('SELECT RE_COMPTES FROM CLEREPAR WHERE RE_AXE = "' + GetControlText('Y_AXE') +
                       '" AND RE_CLE = "' + GetControlText('S_CLEREPARTITION') + '"', True);
  if not Q.EOF then
    SetControlText('GENERECEPT', Q.Fields[0].AsString);
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.Y_AXEChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Char;
begin
  {Récupération du nouvel axe en cours}
  if (GetControlText('Y_AXE') <> '') and (Length(GetControlText('Y_AXE')) > 1) then
    n := string(GetControlText('Y_AXE'))[2]
  else
    n := '0';

  {Si l'axe est compris entre 1 et 5 ...}
  if (n in ['1', '2', '3', '4', '5']) then begin
    {... On met la tablette idoine sur la clef de répartition ...}
    SetControlProperty('S_CLEREPARTITION', 'DATATYPE', 'TTCLEREPART' + n);
    {... et on mémorise l'axe}
    NumAxe := StrToInt(n);
  end
  else
    NumAxe := -1;

  {Si on ne gère pas le CroisAxe}
  if not VH^.AnaCroisaxe then begin
    if GetControlText('Y_JOURNAL') <> AxeJal then begin
      AxeJal := '';
      SetControlText('Y_JOURNAL', '');
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.Y_JOURNALChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  ControleJournal;
end;

{Validation de la saisie avant le lancement du traitement
{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.ControleOK : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  {Mise à jour de la date génération et du journal}
  if not ControlDesZones then Exit;
  
  if VH^.Cpta[AxeToFb(GetControlText('Y_AXE'))].AxGenAttente = '' then
    HShowmessage('18;' + Ecran.Caption + ';Transfert impossible : vous devez définir un compte général d''attente de l''axe.;W;O;O;O;', '', '')
  else if GetControlText('S_CLEREPARTITION') = '' then
    HShowmessage('5;' + Ecran.Caption + ';Vous devez renseigner une clé de répartition.;W;O;O;O;', '', '')
  else if GetControlText('Y_JOURNAL') = '' then
    HShowmessage('6;' + Ecran.Caption + ';Vous devez renseigner un journal.;W;O;O;O;', '', '')
  else if GetControlText('Y_AXE') = '' then
    HShowmessage('7;' + Ecran.Caption + ';Vous devez renseigner un axe analytique.;W;O;O;O;', '', '')
  else if NumAxe = -1 then
    HShowmessage('40;' + Ecran.Caption + ';L''axe analytique n''a pu être lu.;W;O;O;O;', '', '')
  else
    Result := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.ConstituteClausesWhere;
{---------------------------------------------------------------------------------------}
var
  St : string;
begin
  St := GetControlText('GENEEMET');
  if St <> '' then begin
    WhereEmet := AnalyseCompte(St, fbGene, False, False);
    WhereEmet := FindEtReplace(WhereEmet, 'G_GENERAL', 'Y_GENERAL', True);
  end;

  St := GetControlText('GENERECEPT');
  if St <> '' then begin
    WhereRecept := AnalyseCompte(St, fbGene, False, False);
    WhereRecept := FindEtReplace(WhereRecept, 'G_GENERAL', 'Y_GENERAL', True);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.FabricSQLAnal(ModeR, QQ, Sect: string; Emet: Boolean): string;
{---------------------------------------------------------------------------------------}
var
  St   : string;
  StW  : string;
  StXP : string;
  StXN : string;
  SAxe : string;
begin
  SAxe := IntToStr(NumAxe);

  if Sect = '' then begin
    StW := ' Y_DATECOMPTABLE >= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE' ))) + '"'
     + ' AND Y_DATECOMPTABLE <= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE_'))) + '"'
     + ' AND S_CLEREPARTITION = "' + GetControlText('S_CLEREPARTITION') + '"';
  end
  else begin
    StW := ' S_SECTION = "' + Sect + '"'
      + ' AND Y_DATECOMPTABLE >= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE1' ))) + '"'
      + ' AND Y_DATECOMPTABLE <= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE1_'))) + '"';
    if ModeR = 'MC' then ModeR := 'MT'; {pas de subdivision pour les sections réceptrices}
  end;

  StW := StW + ' AND (Y_DEBIT + Y_CREDIT) <> 0 AND Y_QUALIFPIECE = "N" ';

  if ((Emet) and (WhereEmet <> '')) then StW := StW + ' AND ' + WhereEmet;
  if ((not Emet) and (WhereRecept <> '')) then StW := StW + ' AND ' + WhereRecept;

  StXN := StrFPoint(-9 * Resolution(V_PGI.OkDecV + 1));
  StXP := StrFPoint( 9 * Resolution(V_PGI.OkDecV + 1));

  {Sur cumul section/période}
  if ModeR = 'MT' then begin
    if VH^.AnaCroisaxe then
      St := 'SELECT S_SECTION, SUM(Y_DEBIT - Y_CREDIT) AS MNT1, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
          + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_SECTION = Y_SOUSPLAN' + SAxe
          + ' WHERE ' + StW + ' GROUP BY S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING SUM(Y_DEBIT - Y_CREDIT) NOT BETWEEN ' + StXN +
            ' AND ' + StXP
    else
      St := 'SELECT S_SECTION, SUM(Y_DEBIT - Y_CREDIT) AS MNT1'
          + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_AXE = Y_AXE AND S_SECTION = Y_SECTION'
          + ' WHERE ' + StW + ' GROUP BY S_SECTION HAVING SUM(Y_DEBIT-Y_CREDIT) NOT BETWEEN ' + StXN + ' AND ' + StXP;

  end

  {Sur cumul section/général/période}
  else if ModeR = 'MC' then begin
    if VH^.AnaCroisaxe then
      St := 'SELECT S_SECTION, SUM(Y_DEBIT - Y_CREDIT) AS MNT1, Y_GENERAL, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
          + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_SECTION = Y_SOUSPLAN' + SAxe
          + ' WHERE ' + StW + ' GROUP BY S_SECTION, Y_GENERAL, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5 HAVING SUM(Y_DEBIT - Y_CREDIT) NOT BETWEEN '
          + StXN + ' AND ' + StXP
    else
      St := 'SELECT S_SECTION, SUM(Y_DEBIT - Y_CREDIT) AS MNT1, Y_GENERAL'
        + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_AXE = Y_AXE AND S_SECTION = Y_SECTION'
        + ' WHERE ' + StW + ' GROUP BY S_SECTION, Y_GENERAL HAVING SUM(Y_DEBIT - Y_CREDIT) NOT BETWEEN ' + StXN + ' AND ' + StXP;

  end

  {Sur quantité 1 section/période}
  else if ModeR = 'QT1' then begin
    if VH^.AnaCroisaxe then
      St := 'SELECT S_SECTION, SUM(Y_DEBIT - Y_CREDIT) AS MNT1, SUM(Y_QTE1 * (Y_DEBIT - Y_CREDIT) / (Y_DEBIT + Y_CREDIT)) AS MNT2, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
        + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_SECTION = Y_SOUSPLAN' + SAxe
        + ' WHERE ' + StW
    else
      St := 'SELECT S_SECTION, SUM(Y_DEBIT - Y_CREDIT) AS MNT1, SUM(Y_QTE1 * (Y_DEBIT - Y_CREDIT) / (Y_DEBIT + Y_CREDIT)) AS MNT2'
        + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_AXE = Y_AXE AND S_SECTION = Y_SECTION'
        + ' WHERE ' + StW;

    if QQ <> '...' then St := St + ' AND Y_QUALIFQTE1="' + QQ + '"';

    if Sect = '' then begin
      if VH^.AnaCroisaxe then
        St := St + ' GROUP BY S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
      else
        St := St + ' GROUP BY S_SECTION';

      St := St + ' HAVING SUM(Y_DEBIT - Y_CREDIT) NOT BETWEEN ' + StXN + ' AND ' + StXP;
    end

    else begin
      if VH^.AnaCroisaxe then
        St := St + ' GROUP BY S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
      else
        St := St + ' GROUP BY S_SECTION';

      St := St + ' HAVING SUM(Y_QTE1 * (Y_DEBIT - Y_CREDIT) / (Y_DEBIT + Y_CREDIT)) NOT BETWEEN ' + StXN + ' AND ' + StXP
    end;
  end

  {Sur quantité 2 section/période}
  else if ModeR = 'QT2' then begin
    if VH^.AnaCroisaxe then
      St := 'SELECT S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5, SUM(Y_DEBIT - Y_CREDIT) AS MNT1, SUM(Y_QTE2 * (Y_DEBIT - Y_CREDIT) / (Y_DEBIT + Y_CREDIT)) AS MNT2'
          + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_SECTION = Y_SOUSPLAN' + SAxe
          + ' WHERE ' + StW
    else
      St := 'SELECT S_SECTION, SUM(Y_DEBIT - Y_CREDIT) AS MNT1, SUM(Y_QTE2 * (Y_DEBIT - Y_CREDIT) / (Y_DEBIT + Y_CREDIT)) AS MNT2'
          + ' FROM SECTION LEFT JOIN ANALYTIQ ON S_AXE = Y_AXE AND S_SECTION = Y_SECTION'
          + ' WHERE ' + StW;

    if QQ <> '...' then St := St + ' AND Y_QUALIFQTE2 = "' + QQ + '"';

    if Sect = '' then begin
      if VH^.AnaCroisaxe then
        St := St + ' GROUP BY S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
      else
        St := St + ' GROUP BY S_SECTION';

      St := St + ' HAVING SUM(Y_DEBIT - Y_CREDIT) NOT BETWEEN ' + StXN + ' AND ' + StXP;

    end

    else begin
      if VH^.AnaCroisaxe then
        St := St + ' GROUP BY S_SECTION, Y_SOUSPLAN1, Y_SOUSPLAN2, Y_SOUSPLAN3, Y_SOUSPLAN3, Y_SOUSPLAN4, Y_SOUSPLAN5'
      else
        St := St + ' GROUP BY S_SECTION';

      {JP 09/06/05 : Je dois avouer que Y_QTE1 me laisse sceptique, mais aussi loin que je remonte, c'est ainsi :
                     j'aurais eu tendance à vouloir mettre Y_QTE2
      St := St + ' HAVING SUM(Y_QTE1 * (Y_DEBIT - Y_CREDIT) / (Y_DEBIT + Y_CREDIT)) NOT BETWEEN ' + StXN + ' AND ' + StXP}
      St := St + ' HAVING SUM(Y_QTE2 * (Y_DEBIT - Y_CREDIT) / (Y_DEBIT + Y_CREDIT)) NOT BETWEEN ' + StXN + ' AND ' + StXP;
    end;
  end;

  Result := St;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.RemplirLaListe(Q : TQuery; ModeR : string; Source : Boolean);
{---------------------------------------------------------------------------------------}
var
  X : P_RAP;
begin
  X := P_RAP.Create;
  X.Section := Q.FindField('S_SECTION').AsString;

  if Source then begin
    if VH^.AnaCroisaxe then begin
      X.Sousplan[1] := Q.FindField('Y_SOUSPLAN1').AsString;
      X.Sousplan[2] := Q.FindField('Y_SOUSPLAN2').AsString;
      X.Sousplan[3] := Q.FindField('Y_SOUSPLAN3').AsString;
      X.Sousplan[4] := Q.FindField('Y_SOUSPLAN4').AsString;
      X.Sousplan[5] := Q.FindField('Y_SOUSPLAN5').AsString;
    end;

    X.Solde := Q.FindField('MNT1').AsFloat;
    if ModeR = 'MC' then
      X.General := Q.FindField('Y_GENERAL').AsString;
  end

  else begin
    if ((ModeR = 'QT1') or (ModeR = 'QT2')) then X.Solde := Q.FindField('MNT2').AsFloat
                                            else X.Solde := Q.FindField('MNT1').AsFloat;
  end;

  if Source then TSource.Add(X)
            else TDest  .Add(X);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.CoherQte(ModeR, QQ : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  StW     : string;
  Section : string;
  Q       : TQuery;
begin
  Result := True;
  {Sortie triviale}
  if GetCheckBoxState('VerifQTE') = cbChecked then Exit;
  if ((ModeR <> 'QT1') and (ModeR <> 'QT2')) then Exit;

  if QQ = '...' then Exit;

  {Section emettrices}
  StW := ' Y_DATECOMPTABLE >= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE' ))) + '"' +
     ' AND Y_DATECOMPTABLE <= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE_'))) + '"' +
     ' AND S_CLEREPARTITION = "' + GetControlText('S_CLEREPARTITION') + '" AND Y_QUALIFPIECE = "N"';

  if ModeR = 'QT1' then StW := StW + ' AND Y_QUALIFQTE1 <> "' + QQ + '" AND Y_QTE1 <> 0'
                   else StW := StW + ' AND Y_QUALIFQTE2 <> "' + QQ + '" AND Y_QTE2 <> 0';

  if VH^.AnaCroisaxe then
    StW := 'SELECT S_SECTION FROM SECTION LEFT JOIN ANALYTIQ ON S_SECTION = Y_SOUSPLAN' + IntToStr(NumAxe) + ' WHERE ' + StW
  else
    StW := 'SELECT S_SECTION FROM SECTION LEFT JOIN ANALYTIQ ON S_AXE = Y_AXE AND S_SECTION = Y_SECTION WHERE ' + StW;

  if ExisteSQL(StW) then begin
    Result := False;
    Exit;
  end;

  {Section receptrices}
  StW := ' Y_DATECOMPTABLE >= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE1' ))) +
    '" AND Y_DATECOMPTABLE <= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE1_'))) +
    '" AND Y_QUALIFPIECE = "N"';

  if ModeR = 'QT1' then StW := StW + ' AND Y_QUALIFQTE1 <> "' + QQ + '" AND Y_QTE1 <> 0'
                   else StW := StW + ' AND Y_QUALIFQTE2 <> "' + QQ + '" AND Y_QTE2 <> 0';

  Q := OpenSQL('SELECT V_SECTION FROM VENTIL WHERE V_NATURE = "CL' + IntToStr(NumAxe) + '" ' +
               'AND V_COMPTE = "' + GetControlText('S_CLEREPARTITION') +
               '" ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL', True);

  if Q.EOF then begin
    Ferme(Q);
    Exit;
  end;

  try
    while not Q.EOF do begin
      Section := Q.FindField('V_SECTION').AsString;
      if Section = '' then Continue;

      if VH^.AnaCroisaxe then
        StW := 'SELECT S_SECTION FROM SECTION LEFT JOIN ANALYTIQ ON S_SECTION = Y_SOUSPLAN' +
               IntToStr(NumAxe) + ' WHERE ' + StW + ' AND S_SECTION = "' + Section + '"'
      else
        StW := 'SELECT S_SECTION FROM SECTION LEFT JOIN ANALYTIQ ON S_AXE = Y_AXE AND S_SECTION = Y_SECTION' +
               ' WHERE ' + StW + ' AND S_SECTION = "' + Section + '"';


      if ExisteSQL(StW) then begin
        Result := False;
        Exit;
      end;
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.MajDernTIS;
{---------------------------------------------------------------------------------------}
var
  Q : TOB;
begin
  Q := TOB.Create('EDTLEGAL', nil, -1);
  Q.SetString('ED_OBLIGATOIRE'  , '-');
  Q.SetString('ED_TYPEEDITION'  , 'TIS');
  Q.SetString('ED_EXERCICE'     , GetControlText('Y_AXE'));
  Q.SetString('ED_UTILISATEUR'  , V_PGI.User);
  Q.SetString('ED_DESTINATION'  , GetControlText('S_CLEREPARTITION'));
  Q.SetDateTime('ED_PERIODE'    , NowH);
  Q.SetDateTime('ED_DATEEDITION', Date);
  Q.SetDateTime('ED_DATE1'      , StrToDate(GetControlText('Y_DATECOMPTABLE' )));
  Q.SetDateTime('ED_DATE2'      , StrToDate(GetControlText('Y_DATECOMPTABLE_')));
  Q.InsertDB(nil);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.DejaTIS : Boolean;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  Result := False;

  SQL := 'SELECT ED_OBLIGATOIRE FROM EDTLEGAL WHERE ED_OBLIGATOIRE = "-" AND ED_TYPEEDITION = "TIS" AND ' +
         'ED_EXERCICE = "' + GetControlText('Y_AXE') + '" AND ED_DESTINATION="' + GetControlText('S_CLEREPARTITION') + '"';
  SQL := SQL + ' AND ((ED_DATE1 <= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE' ))) +
                '" AND ED_DATE2 >= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE' ))) + '")' +
                 ' OR (ED_DATE1 <= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE_'))) +
                '" AND ED_DATE2 >= "' + USDateTime(StrToDate(GetControlText('Y_DATECOMPTABLE_'))) + '"))';

  if ExisteSQL(SQL) then
    Result := HShowmessage('17;' + Ecran.Caption + ';Un transfert a déjà été effectué sur tout ou partie de la période choisie.'#13 +
                           'Confirmez-vous l''opération ?;Q;YN;Y;Y;', '', '') <> mrYes;
end;



{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.GestionFlash(Fin : Boolean = True);
{---------------------------------------------------------------------------------------}
begin
  SetControlVisible('FLASHVIDE', not Fin);
  if fin then SourisNormale
         else SourisSablier;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.ChargeDest(ModeR, QQ : string) : string;
{---------------------------------------------------------------------------------------}
var
  QV     : TQuery;
  QDest  : TQuery;
  St     : string;
  Section: string;
begin
  QV := OpenSQL('SELECT V_SECTION FROM VENTIL WHERE V_NATURE = "CL' + IntToStr(NumAxe) + '" AND ' +
                'V_COMPTE = "' + GetControlText('S_CLEREPARTITION') + '" ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL', True);

  if QV.EOF then begin
    Ferme(QV);
    Exit;
  end;

  InitMove(QV.RecordCount, 'Chargement des infos dest.');
  try
    while not QV.EOF do begin
      MoveCur(False);
      Section := QV.FindField('V_SECTION').AsString;
      if Section = '' then Continue;
      St := FabricSQLAnal(ModeR, QQ, Section, False);
      {La requête n'est peut être pas correcte vu le nombre de concaténations dont elle est issues}
      try
        QDest := OpenSQL(St, True);
      except
        on E : Exception do begin
          MessageAlerte(E.Message + #13#13 + 'Vérifiez la définition des critères sur les comptes généraux.');
          Exit;
        end;
      end;

      while not QDest.EOF do begin
        RemplirLaListe(QDest, ModeR, False);
        QDest.Next;
      end;
      Ferme(QDest);
      QV.Next;
    end;
  finally
    Ferme(QV);
    FiniMove;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.GenereLesPieces;
{---------------------------------------------------------------------------------------}
var
  isource, idest, NumV: Integer;
  Facturier: string;
  Q : TQuery;
  NewNum: Longint;
  Total, Montant, Debit, Credit, TotDest : Double;
  Section, SectDest, General, GeneDest: String17;
  X, Z: P_RAP;
  T : TOB;
begin
  InitMove(TSource.Count, ' ');
  Q := OpenSQL('SELECT J_COMPTEURNORMAL FROM JOURNAL WHERE J_JOURNAL = "' + GetControlText('Y_JOURNAL') + '"', True);
  if not Q.EOF then
    Facturier := Q.FindField('J_COMPTEURNORMAL').AsString;
  Ferme(Q);

  try
    TotDest := 0;
    for iDest := 0 to TDest.Count - 1 do begin
      Z := P_RAP(TDest[iDest]);
      TotDest := TotDest + Z.Solde;
    end;

    for isource := 0 to TSource.Count - 1 do begin
      MoveCur(False);

      SetIncNum(EcrAna, Facturier, NewNum, DateGene);
      X := P_RAP(TSource[isource]);

      Section := X.Section;
      General := X.General;
      Debit   := 0;
      Credit  := 0;

      if VH^.MontantNegatif then begin
        if X.Solde < 0 then Credit := X.Solde
                       else Debit  := -X.Solde;
      end
      else begin
        if X.Solde > 0 then Credit := X.Solde
                       else Debit  := -X.Solde;
      end;
      NumV := 1;

      T := CreerLigneAna(NewNum, NumV, Section, General, Debit, Credit, X.Sousplan);
      MajSoldeSectionTOB(T, True);

      Total := 0;
      for iDest := 0 to TDest.Count - 1 do begin
        Z := P_RAP(TDest[iDest]);
        SectDest := Z.Section;
        GeneDest := Z.General;

        Debit := 0;
        Credit := 0;
        if ((iDest < TDest.Count - 1) and (TotDest <> 0)) then begin
          Montant := Arrondi(X.Solde * Z.Solde / TotDest, V_PGI.OkDecV);
          Total := Total + Montant;
        end
        else
          Montant := Arrondi(X.Solde - Total, V_PGI.OkDecV);

        if Montant = 0 then Continue;

        if X.Solde > 0 then begin
          if ((Montant > 0) or (VH^.MontantNegatif)) then Debit  := Montant
                                                     else Credit := Abs(Montant);
        end

        else begin
          if ((Montant < 0) or (VH^.MontantNegatif)) then Credit := - Montant
                                                     else Debit  := Montant;
        end;
        Inc(NumV);
        T := CreerLigneAna(NewNum, NumV, SectDest, General, Debit, Credit, X.Sousplan); {gene = celui de origine}
        MajSoldeSectionTOB(T, True);
      end;
    end;
  finally
    FiniMove;
  end;
end;

{JP 10/06/05 : J'ai quelques interrogations concernant les codes ci-dessous c'est ainsi au moins
               depuis la V5 : Y_NumLigne = 0 !!
                              Y_DEBIT = Y_DEBITDEV !!
                              Y_CREDIT = Y_CREDITDEV !!
                              les quantités ne sont pas gérées ??
{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.CreerLigneAna(NewNum, NumV : Integer; Section, General: string;
                                  Debit, Credit : Double; SousPlan: array of string) : TOB;
{---------------------------------------------------------------------------------------}
var
  premier_axe : Byte;
  i : Integer;
  T : TOB;
begin
  T := TOB.Create('ANALYTIQ', vTobAna, -1);
  premier_axe := RecherchePremDerAxeVentil.premier_axe;
  if General = '' then
    General := VH^.Cpta[AxeToFb(GetControlText('Y_AXE'))].AxGenAttente;
  T.PutValue('Y_GENERAL', General);

  if VH^.AnaCroisaxe then T.PutValue('Y_AXE', 'A' + IntToStr(premier_axe))
                     else T.PutValue('Y_AXE', GetControlText('Y_AXE'));

  if NumV = 1 then T.PutValue('Y_TYPEMVT', 'AE')
              else T.PutValue('Y_TYPEMVT', 'AL');
  T.PutValue('Y_DATECOMPTABLE' , DateGene);
  T.PutValue('Y_PERIODE'       , GetPeriode(DateGene));
  T.PutValue('Y_SEMAINE'       , NumSemaine(DateGene));
  T.PutValue('Y_NUMLIGNE'      , 0);
  T.PutValue('Y_NUMEROPIECE'   , NewNum);
  T.PutValue('Y_EXERCICE'      , QuelExo(GetControlText('DATEGENERATION')));
  T.PutValue('Y_REFINTERNE'    , Copy(GetControlText('Y_REFINTERNE'), 1, 35));
  T.PutValue('Y_LIBELLE'       , Copy(GetControlText('Y_LIBELLE'), 1, 35));
  T.PutValue('Y_QUALIFPIECE'   , 'N');
  T.PutValue('Y_NATUREPIECE'   , 'OD');
  T.PutValue('Y_DEBIT'         , Debit);
  T.PutValue('Y_CREDIT'        , Credit);
  T.PutValue('Y_DEBITDEV'      , Debit);          ////
  T.PutValue('Y_CREDITDEV'     , Credit);         ////
  T.PutValue('Y_ETABLISSEMENT' , VH^.EtablisDefaut);
  T.PutValue('Y_DEVISE'        , V_PGI.DevisePivot);
  T.PutValue('Y_TAUXDEV'       , 1);
  T.PutValue('Y_TYPEANALYTIQUE', 'X');
  T.PutValue('Y_DATEAUXDEV'    , DateGene);
  T.PutValue('Y_TOTALECRITURE' , 0);
  T.PutValue('Y_TOTALDEVISE'   , 0);
  T.PutValue('Y_JOURNAL'       , GetControlText('Y_JOURNAL'));
  T.PutValue('Y_NUMVENTIL'     , NumV);
  T.PutValue('Y_CONTROLE'      , '-');
  T.PutValue('Y_ECRANNOUVEAU'  , 'N');
  T.PutValue('Y_CREERPAR'      , 'GEN');
  T.PutValue('Y_EXPORTE'       , '---');
  T.PutValue('Y_QUALIFCRQTE1'  , '...');
  T.PutValue('Y_QUALIFCRQTE2'  , '...');

  {Gestion Ana croisaxe}
  T.PutValue('Y_SOUSPLAN' + IntToStr(NumAxe), Section );
  if VH^.AnaCroisaxe then begin
    for i := 1 to MaxAxe do begin
      if i = NumAxe then Continue;
      T.PutValue('Y_SOUSPLAN' + IntToStr(i), SousPlan[i-1] );
    end;
    T.PutValue('Y_SECTION', T.GetString('Y_SOUSPLAN' + IntToStr(premier_axe)) );
  end

  else begin
    for i := 1 to MaxAxe do
      T.PutValue('Y_SOUSPLAN' + IntToStr(i), '');
    T.PutValue('Y_SECTION', Section);
  end;

  T.InsertDB(nil);

  if ((NumV = 1) and (V_PGI.IoError = oeOk)) then
    TPieces.Add(T);
  Result := T;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVIDESECTIONS.ViderListe(Detruire : Boolean = False);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TSource) then begin
    VideListe(TSource);
    if Detruire then FreeAndNil(TSource);
  end;

  if Assigned(TDest) then begin
    VideListe(TDest);
    if Detruire then FreeAndNil(TDest);
  end;

  {Il vaut mieux vider vTobAna avant TPieces qui contient les filles de vTobAna}
  if Assigned(TPieces) then begin
    VideListe(TPieces);
    if Detruire then FreeAndNil(TPieces);
  end;

  if Assigned(vTobAna) then begin
    vTobAna.ClearDetail;
    if Detruire then FreeAndNil(vTobAna);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.ControleJournal : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  J : string;
begin
  Result := True;
  AxeJal := '';
  J := GetControlText('Y_JOURNAL');
  if J = '' then Exit;
  {Contrôles sur le journal sélectionné}
  Q := OpenSQL('SELECT J_AXE, J_COMPTEURNORMAL FROM JOURNAL WHERE J_JOURNAL = "' + J + '"', True);
  try
    if not Q.EOF then begin
      AxeJal := Q.Fields[0].AsString;
      {Contrôle sur l'axe analytique du journal}
      if AxeJal <> GetControlText('Y_AXE') then begin
        SetControlText('Y_JOURNAL', '');
        AxeJal := '';
        HShowmessage('15;' + Ecran.Caption + ';Vous devez renseigner un journal qui porte sur l''axe choisi.;W;O;O;O;', '', '');
        Result := False;
      end
      {Contrôle sur le compteur du journal}
      else if Q.Fields[1].AsString = '' then begin
        SetControlText('Y_JOURNAL', '');
        AxeJal := '';
        HShowmessage('19;' + Ecran.Caption + ';Vous devez renseigner un journal qui possède un facturier.;W;O;O;O;', '', '');
        Result := False;
      end;
    end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.ControleDate : Boolean;
{---------------------------------------------------------------------------------------}
var
  DD: TDateTime;
  Err: integer;
begin
  {Contrôle du format de la date}
  if not IsValidDate(GetControlText('DATEGENERATION')) then begin
    HShowmessage('0;' + Ecran.Caption + ';Vous devez renseigner une date valide.;W;O;O;O;', '', '');
    SetControlText('DATEGENERATION', DateToStr(V_PGI.DateEntree));
    DateGene := V_PGI.DateEntree;
    Result   := False;
  end

  {Contrôle de la validité de la date}
  else begin
    DD := StrToDate(GetControlText('DATEGENERATION'));
    Err := DateCorrecte(DD);
    Result := Err < 1;
    case Err of
      1, 2 : HShowmessage('1;' + Ecran.Caption + ';La date que vous avez renseignée est sur un exercice non ouvert.;W;O;O;O;', '', '');
      3    : HShowmessage('3;' + Ecran.Caption + ';La date que vous avez renseignée est antérieure à la clôture provisoire.;W;O;O;O;', '', '');
      else begin
        if RevisionActive(DD) then begin
          SetControlText('DATEGENERATION', DateToStr(V_PGI.DateEntree));
          DateGene := V_PGI.DateEntree;
        end
        else
          DateGene := DD;
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPVIDESECTIONS.ControlDesZones : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ControleJournal and ControleDate;
end;

initialization
  RegisterClasses([TOF_CPVIDESECTIONS]);

end.

