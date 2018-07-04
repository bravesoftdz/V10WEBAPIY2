unit dpTOFAnnuLien;
// TOF de toutes les fiches DP LIENINTERxxxxxx

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, HTB97,
  HEnt1, hmsgbox, HDB,
{$IFDEF EAGLCLIENT}
  MaineAGL, eMul, UTob, utilEagl,
{$ELSE}
  FE_Main, Mul, {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,EdtrEtat,{$ENDIF}
{$ENDIF}
  HQry, Menus,
{$IFDEF VER150}
  Variants,
{$ENDIF}
  UTOF, AGLInit;

 type
  TOF_ANNULIEN = class(TOF)
    MenuImpression : TPopUpMenu;
  public
    procedure OnArgument(Arguments : String ) ; override ;
    procedure OnLoad ; override;
    procedure OnClose ; override ;
  private
    Formjur      : String;
    GuidPerDos   : String;
    Fct          : String;
    GrpFiscal    : String;
    Count        : integer;
    procedure BDelete_OnClick(Sender: TObject) ;
    procedure BImprimer_OnClick(Sender: TObject) ;
    procedure BTLIENANNU_OnClick(Sender : TObject);
    procedure OnDblClickFListe(Sender : TObject);
    procedure BTANNUAIRE_OnClick(Sender : TObject);
    procedure TraiteListeFonctions(fct1: String=''; fct2: String='');
    procedure OnClickPopFonction(Sender : TObject);
    function  ListeCorrecte(nomchp, titrechp: String): Boolean;
    procedure PurgePopup;
    procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure MenuImpressionAnnuaireFonction_OnClick (Sender : TObject);
    procedure MenuImpressonAnnuaireAlphabetique_OnClick (Sender : TObject);

{$IFDEF VEGA}
    procedure OnVegaClickOuvre   (Sender:TObject);
{$ENDIF}
  end;


///////////// IMPLEMENTATION ////////////
implementation

uses galOutil, UtilLiensAnnuaire, AnnOutils,
     dpOutils, DpJurOutils, UtilMulTraitmt;

     
procedure TOF_ANNULIEN.OnClose;
begin
  // utile uniquement à certaines fiches...
  if (Ecran.Name='DETLIENJURIDIQUE') or (Ecran.Name='LIENINTERPROP') or (Ecran.Name='LIENINTERVENANT') then
  begin
          // Ajout CM 02/05/06 Pour Vega, besoin du guidper en retour
{$IFDEF VEGA}
          if (Ecran.ModalResult = mrOk) and (VarToStr (GetField ('ANL_GUIDPER')) <> '') then
              TFMul (Ecran).Retour := 'ANN_GUIDPER=' + VarToStr (GetField ('ANL_GUIDPER'))
                              + ';ANL_GUIDPERDOS=' + VarToStr (GetField ('ANL_GUIDPERDOS'))
                              + ';ANL_TYPEDOS=' + VarToStr (GetField ('ANL_TYPEDOS'))
                              + ';ANL_NOORDRE=' + VarToStr (GetField ('ANL_NOORDRE'))
                              + ';ANL_FONCTION=' + VarToStr (GetField ('ANL_FONCTION'))
          else
              TFMul (Ecran).Retour := '';
{$ELSE}
          // sortie interdite si manque l'info !
          if not ListeCorrecte ('ANL_NOMPER', 'Nom abrégé de la personne') then
             exit;
          TFMul (Ecran).Retour := VarToStr (GetField ('ANL_NOMPER'));
{$ENDIF}
  end
  // mise à jour des filiales => il faut recalculer les éléments
  else if (Ecran.Name='LIENINTERFILIALE') and (VarToStr (GetField ('ANL_GUIDPER')) <> '') then
  begin
      MajNbFilialesOrga( VarToStr (GetField ('ANL_GUIDPER')) );
  end;

  PurgePopup;
  inherited;
end;

{$IFDEF VEGA}
procedure TOF_ANNULIEN.OnVegaClickOuvre (Sender:TObject);
begin
     Ecran.ModalResult := mrOk;
end;
{$ENDIF}

procedure TOF_ANNULIEN.OnArgument(Arguments : String ) ;
// contient Guidperdos + ;DOS|JUR|Fnc avec Fnc = la fonction voulue (famper)
var Q1    : TQuery;
    Vales : HTStrings;
    cbfnc : THValComboBox;
    sRequete : String;
begin
  inherited;

  SetControlVisible( 'PCOMPLEMENT', MontreOnglet( TFMul(Ecran), 'PCOMPLEMENT' ));
  // ne pas passer ACTION= dans les arguments... (non traité)
  GuidPerDos := ReadTokenSt(Arguments);
  Fct := ReadTokenSt(Arguments);
  GrpFiscal := ReadTokenSt(Arguments);
  if GrpFiscal <> 'X' then GrpFiscal := '-';
  // Cette zone restait active car les fiches étaient souvent appelées avec
  // ANL_FORME= quelque chose, mais beaucoup de liens n'ont pas ANL_FORME renseigné,
  // donc il fallait pouvoir revenir à <<Tous>> dans cette combo
  // Rq : seuls les fonctions pour lesquelles JFT_DEFAUT est à 'X' doivent avoir
  // ANL_FORME renseigné...
  // Par contre, depuis que les appels aux fiches ne contiennent plus ANL_FORME=,
  // on peut empécher de choisir une forme dans la zone (car aucun intérêt pour des liens)
  SetControlEnabled('ANL_FORME', False);

  if GuidPerDos<>'' then
    begin
    // Récupère la forme juridique, car elle n'est pas passée dans Argument
    // mais plus souvent dans Range du AglLanceFiche...
    Q1 := OpenSQL('select ANN_FORME from ANNUAIRE where ANN_GUIDPER = "'+ GuidPerDos+'"', TRUE);
    if not Q1.eof then Formjur := Q1.Fields[0].AsString;
    Ferme(Q1);
    end;

  // Droits de création,
  GereDroitsConceptsAnnuaire(Self, Ecran);

  // appel Liens personne depuis la sélection de dossier (clic-droit)
  if Fct = 'DOS' then
    // (donc en fait équivaut à Ecran.Name='LIENPERSONNE' or Ecran.Name='LIENINTERDOS')
    begin
    // on veut juste visualiser dans le mul, mais pas de modifs
    SetControlVisible ('BTLIENANNU', FALSE);
    SetControlVisible ('BTANNUAIRE', FALSE);
    SetControlVisible ('BINSERT', FALSE);
    SetControlVisible ('BDELETE', FALSE);
{ 5/07/02 => si il reste 2 liens (ex: ASS/INT) ET d'autres liens, le bouton
  delete était caché donc impossibilité de supprimer le lien INT...
  Marchait que si il ne restait plus qu'une paire de liens ASS/INT  !
    SetControlVisible ('BDELETE', FALSE);
    // Autorise la suppression du lien INTERVENANT (loi NRE)
    // s'il est tout seul et non lié à un dossier juridique
    if Not ExisteSQL('select ANL_GUIDPER from ANNULIEN where ANL_GUIDPERDOS="'+GuidPerDos+'"'
     +' AND ANL_FONCTION<>"INT"') // pas de lien autre que intervenant
    // et vérifie que l'éventuel lien INT n'est pas utilisé par un dossier juridique
    and Not ExisteSQL('select ANL_GUIDPER from ANNULIEN where ANL_GUIDPERDOS="'+GuidPerDos+'"'
       +' AND ANL_FONCTION="INT" AND ANL_CODEDOS<>"&#@"') then
      SetControlVisible ('BDELETE', True);
}
    end
  // appel depuis l'annuaire juridique
  else if Fct = 'JUR' then
    begin
    // choix du type de lien dans un popup (voir fiche DP LIENINTERJURI)
    PurgePopup;
    // création popup
    TraiteListeFonctions;
    end
{  // BTCAC depuis fiche juridique
  else if Fct = 'CCT' then
    begin
    // choix du type de lien dans un popup (voir fiche DP LIENINTERJURI)
    PurgePopup;
    // création popup
    TraiteListeFonctions('CCT', 'CCS');
    end
}  // autres liens : pas de déploiement d'un popup mais
  // le bouton ouvrira directement les liens
  else
    TToolbarButton97(GetControl('BTLIENANNU')).Onclick := BTLIENANNU_OnClick;

  //####=> pour toutes les fiches dédiées à une fonction, on devrait mettre
  // directement la valeur
  if Fct='FIL' then // ou Ecran.Name='LIENINTERFILIALE', 'LIENTETEGROUPE'
    SetControlText('ANL_FONCTION','FIL')
  else if Fct='OGA' then // ou LIENINTERCGE ou LIENINTERVENANT
    begin
    SetControlText('ANL_FONCTION', 'OGA');
    if Ecran.Name='LIENINTERCGE' then
      THValComboBox(GetControl('ANL_TYPEPER')).plus := 'JTP_FAMPER="OGA"';
    end
  // Commissaire titulaire ou commissaire suppléant
  else if Fct = 'CCT' then
    SetControlProperty('ANL_FONCTION','Plus', 'AND (JTF_FONCTION = "CCT" OR JTF_FONCTION = "CCS")')
  // autres liens de type juridique => on limite les fonctions
  // à celles disponibles pour la forme jur en cours (c'est mieux que rien !)
  else if ExisteSQL('select JFT_FONCTION from JUFONCTION where JFT_FONCTION="'+Fct+'"'
   +' and JFT_TYPEDOS="STE" and JFT_FONCTION <> "INT"') then //JFT_TIERS<>"X"') then
   begin
   cbfnc := THValComboBox(GetControl('ANL_FONCTION'));
   cbfnc.DataType := ''; // car tablette JUTYPEFONCT ne tient pas cpte de la forme jur
   cbfnc.Items.Clear;
   cbfnc.Values.Clear;
   sRequete := 'select JTF_FONCTION, JTF_FONCTABREGE'
    +' from JUTYPEFONCT left join JUFONCTION on JFT_FONCTION=JTF_FONCTION'
    +' where JFT_FORME="'+Formjur+'" AND JFT_TYPEDOS="STE" AND JFT_FONCTION <> "INT"'//JFT_TIERS<>"X"'
    +' order by JFT_TRI';
   Q1 := OpenSQL(sRequete,true, -1, '', True);
   Vales := HTStringList.Create;
   cbfnc.Items.Add('<<Tous>>');
   Vales.Add('') ;
   While not Q1.EOF do
     begin
     Vales.Add(Q1.FindField('JTF_FONCTION').AsString) ;
     cbfnc.Items.Add(Q1.FindField('JTF_FONCTABREGE').AsString);
     Q1.Next;
     end;
   THValComboBox(GetControl('ANL_FONCTION')).Values := Vales;
   Vales.Free;
   Ferme(Q1);
   end;

  // certaines fiches permettait le selectall, mais pas les enreg séparés !
{$IFDEF EAGLCLIENT}
  THDBGrid(GetControl('FListe')).MultiSelect := True;
{$ELSE}
  THDBGrid(GetControl('FListe')).MultiSelection := True;
{$ENDIF}
  THDBGrid(GetControl('FListe')).OnDblClick := OnDblClickFListe;
  TButton(GetControl('BTANNUAIRE')).OnClick := BTANNUAIRE_OnClick;
  Ecran.OnKeyDown := Form_OnKeyDown;

  // pour gérer suppression du lien INT si il est le dernier restant
  TToolbarButton97(GetControl('BDelete')).OnClick := BDelete_OnClick;

  //--- CAT le 14/02/2008 Fiche qualitée : 11993 -> Edition de la liste des filiales
  TToolbarButton97(GetControl('BImprimer')).OnClick := BImprimer_OnClick;
  //--- CAT le 14/02/2008 Fiche qualitée : 11992 -> Edition de l'annuaire spécifique du dossier
  if (Ecran.Name='LIENINTERDOS') and (GetControl('IMP_MENU')<>Nil) then
   begin
    MenuImpression := TPopupMenu(GetControl('IMP_MENU'));
    TToolbarButton97(GetControl('BImprimer')).DropDownMenu:=MenuImpression;
    TToolbarButton97(GetControl('BImprimer')).Width:=38;
    TMenuItem(GetControl('IMPANNFONC')).OnClick:=MenuImpressionAnnuaireFonction_OnClick;
    TMenuItem(GetControl('IMPANNALPHA')).OnClick:=MenuImpressonAnnuaireAlphabetique_OnClick;
   end;

{$IFDEF VEGA}
  // $$$ JP 03/07/06: bouton ouvrir pour la sélection de la personne
  TToolBarButton97 (GetControl ('BOuvrir')).OnClick := OnVegaClickOuvre;
{$ENDIF}
end;


procedure TOF_ANNULIEN.OnLoad;
begin
  if Fct = 'SAP' then
    Ecran.Caption := 'SOCIETE APPARENTE '
  else if Fct = 'SOC' then
    Ecran.Caption := 'Organismes sociaux rattachés au dossier '
  else if Fct = 'FIS' then
    Ecran.Caption := 'Organismes fiscaux rattachés au dossier '
  else if Fct = 'TIF' then
    Ecran.Caption := 'SOCIETE-MERE '
  else if Fct = 'CGE' then
    Ecran.Caption := 'Centre de gestion rattaché au dossier ';
  if (Fct='SAP') or (Fct='SOC') or (Fct='FIS')
   or (Fct='TIF') or (Fct='CGE') then
    UpdateCaption(Ecran);
end;


procedure TOF_ANNULIEN.BDelete_OnClick(Sender: TObject);
var
    Lst: THDBGrid;
    SGuiddos, typedos, fonc, Guidper : String;
    i, nb : Integer;
    Q : THQuery;
begin
  // inherited inutile car ne rétablit pas le fonctionnement du bouton delete
  Lst := TFMul(Ecran).FListe;
  Q := TFMul(Ecran).Q;
  // vérifs
  if GuidPerDos='' then exit;
  if (Not Lst.AllSelected) and (Lst.NbSelected<1) then
    begin PGIInfo('Aucun élément sélectionné.', TitreHalley); exit; end;
  If Not ListeCorrecte('ANL_TYPEDOS', 'Type de dossier') then exit;
  If Not ListeCorrecte('ANL_FONCTION', 'Nature de lien') then exit;
  If Not ListeCorrecte('ANL_GUIDPERDOS', 'Numéro') then exit;
  // If Not ListeCorrecte('ANL_FORME', 'Forme du dossier') then exit;
  if PGIAsk('Confirmez-vous la suppression du(des) lien(s) ?', TitreHalley)=mrNo then exit;

  // traitement
  if Lst.AllSelected then
    BEGIN
{$IFDEF EAGLCLIENT}
    if not TFMul(Ecran).FetchLesTous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
      begin
      // le query contient tous les enreg
      Q.First;
      while Not Q.EOF do
        begin
        // GuidPerDos est parfois utilisé en critère (Sélection de dossier / clic-droit /
        // liens annuaire, ou Annuaire / Clic-droit / Liens => ouvre le mul avec
        // anl_Guidper=Guidperdosreçuenarg and anl_Guidperdos = ann_Guidper, donc liste
        // des dossiers sur lesquels la fiche en cours INTERVIENT)
        // donc obligé de le récupérer :
        SGuiddos := Q.FindField('ANL_GUIDPERDOS').AsString;
        typedos := Q.FindField('ANL_TYPEDOS').AsString;
        fonc := Q.FindField('ANL_FONCTION').AsString;
        Guidper := Q.FindField('ANL_GUIDPER').AsString;
        if Guidper<>'' then
        begin
          SupprimeLienAnnuaire(SGuiddos, Guidper, typedos, fonc);

          {+GHA 11/2007 - FQ 11860.}
          if fonc = 'TIF' then
            ModifieLienAnnuaire(Guidper, SGuiddos, typedos, 'FIL');
          {-GHA 11/2007 - FQ 11860.}
        end;
        Q.Next;
        end;
      end;
    END
  else
  BEGIN
    nb := Lst.NbSelected;
    for i:=0 to nb-1 do
    begin
      Lst.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(Lst.Row - 1) ;
{$ENDIF}
      SGuiddos := Q.FindField('ANL_GUIDPERDOS').AsString;
      typedos := Q.FindField('ANL_TYPEDOS').AsString;
      fonc := Q.FindField('ANL_FONCTION').AsString;
      Guidper := Q.FindField('ANL_GUIDPER').AsString;
      if Guidper <> '' then
      begin
        SupprimeLienAnnuaire(SGuiddos, Guidper, typedos, fonc);
        {+GHA 11/2007 - FQ 11860.}
        if fonc = 'TIF' then
          ModifieLienAnnuaire(Guidper, SGuiddos, typedos, 'FIL');
        {-GHA 11/2007 - FQ 11860.}
      end;
    end;
  END;
  // désélectionne
  FinTraitmtMul(TFMul(Ecran));
  // actualisation de l'écran
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
  // MD20070709 FQ11590
  if Ecran.Name='LIENINTERFILIALE' then MajNbFilialesOrga( GuidPerDos );
end;

procedure TOF_ANNULIEN.BImprimer_OnClick(Sender: TObject);
var ChSql : String;
begin
 if Ecran.Name='LIENINTERFILIALE' then
  begin
   ChSql:='SELECT * from ANNULIEN '+
          'LEFT JOIN ANNUAIRE ON ANN_GUIDPER=ANL_GUIDPER '+
          'LEFT JOIN DOSSIER ON DOS_GUIDPER=ANL_GUIDPER '+
          'WHERE ANL_GUIDPERDOS="'+GuidPerDos+'" '+
          'AND ANL_FONCTION="FIL" '+
          'ORDER BY ANL_NOMPER';
   LanceEtat ('E','DPG','FP1',True,False,False,nil,ChSql,'Documents permanents',False);
  end
 else
  if (Ecran.Name<>'LIENINTERDOS') then
   TFMul(Ecran).BImprimerClick(Sender);
end;

procedure TOF_ANNULIEN.MenuImpressionAnnuaireFonction_OnClick (Sender : ToBject);
var ChSql : String;
begin
 ChSql:='SELECT *, LEFT(ANL_NOMPER,1) RUPT FROM ANNULIEN '+
        'LEFT JOIN ANNUAIRE ON ANN_GUIDPER=ANL_GUIDPER '+
        'LEFT JOIN DOSSIER ON DOS_GUIDPER=ANL_GUIDPERDOS '+
        'LEFT JOIN JUTYPEFONCT ON JTF_FONCTION=ANL_FONCTION '+
        'WHERE ANL_GUIDPERDOS="'+GuidPerDos+'" '+
        'AND ANL_FONCTION<>"INT" '+
        'AND ANL_FONCTION<>"STE" '+
        'AND ANL_FONCTION<>"TRS" '+
        'ORDER BY ANL_NOMPER,ANL_TYPEPER';

 LanceEtat ('E','DPG','AN1',True,False,False,nil,ChSql,'Annuaire par fonction',False);
end;

procedure TOF_ANNULIEN.MenuImpressonAnnuaireAlphabetique_OnClick (Sender : TObject);
var ChSql : string;
begin
 ChSql:='SELECT * FROM ANNULIEN '+
        'LEFT JOIN ANNUAIRE ON ANN_GUIDPER=ANL_GUIDPER '+
        'LEFT JOIN DOSSIER ON DOS_GUIDPER=ANL_GUIDPERDOS '+
        'LEFT JOIN JUTYPEFONCT ON JTF_FONCTION=ANL_FONCTION '+
        'WHERE ANL_GUIDPERDOS="'+GuidPerDos+'" '+
        'AND ANL_FONCTION<>"INT" '+
        'AND ANL_FONCTION<>"STE" '+
        'AND ANL_FONCTION<>"TRS" '+
        'ORDER BY ANL_FONCTION,ANL_TYPEPER';

 LanceEtat ('E','DPG','AN2',True,False,False,nil,ChSql,'Annuaire alphabétique',False);
end;


procedure TOF_ANNULIEN.TraiteListeFonctions(fct1: String=''; fct2: String='');
// appelée dans le cas de l'annuaire juridique
var
  Q1 : TQuery;
  sLaFonc, sRequete, LibAbreg : String;
  FncItem : TMenuItem;
  Vales : HTStrings;
begin
  sRequete := 'select JTF_FONCTION, JTF_FONCTABREGE'
   +' from JUTYPEFONCT left join JUFONCTION on JFT_FONCTION=JTF_FONCTION'
   +' where JFT_FORME="'+Formjur+'" and JFT_TYPEDOS="STE" '
   +'   and JFT_FONCTION <> "INT" AND (JTF_FONCTION <> "STE") ';
   //JFT_TIERS<>"X"';
  if (fct1<>'') and (fct2<>'') then
    sRequete := sRequete + ' and (JFT_FONCTION="'+fct1+'" or JFT_FONCTION="'+fct2+'")'
  else
    begin
    if fct1<>'' then sRequete := sRequete + ' and JFT_FONCTION="'+fct1+'"';
    if fct2<>'' then sRequete := sRequete + ' and JFT_FONCTION="'+fct2+'"';
    end;
  sRequete := sRequete + ' order by JFT_TRI';
  Q1 := OpenSQL(sRequete, true, -1, '', True);
  Count := 0;
  Vales := HTStringList.Create;
  THValComboBox(GetControl('ANL_FONCTION')).Items.Add('<<Tous>>');
  Vales.Add('') ;
  while not Q1.EOF do
    begin
    LibAbreg := Q1.FindField('JTF_FONCTABREGE').AsString;
    sLaFonc  := Q1.FindField('JTF_FONCTION').AsString;
    FncItem := TMenuItem.Create(GetControl('POPMENUJUR'));
    FncItem.Caption := LibAbreg;
    FncItem.Hint := sLaFonc;
    Vales.Add(sLaFonc) ;
    THValComboBox(GetControl('ANL_FONCTION')).Items.Add(LibAbreg);
    FncItem.OnClick := OnClickPopFonction;
    Inc(Count);
    TPopupMenu(GetControl('POPMENUJUR')).Items.Add(FncItem);
    Q1.Next;
    end;
  THValComboBox(GetControl('ANL_FONCTION')).Values := Vales;
  Vales.Free;
  Ferme(Q1);
end;


procedure TOF_ANNULIEN.OnClickPopFonction(Sender : TObject);
var
  Cle,GuidPer,Lafct : String;
  MI : TMenuItem;
begin
  MI := Sender as TMenuItem;
  Lafct := MI.Hint;
  // tous les liens juridiques se font avec STE (chp ANL_TYPEDOS)
  // choix d'une personne dans le mul lienannuaire
  GuidPer := AGLLanceFiche('DP','LIENANNUAIRE','GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafct,
    '', ''); // #### on ne passe pas LaFct dans Arguments => voir traitmt dans dpTofAnnuaire
  // si personne choisie, on renseigne sa fiche d'intervention juridique
  if GuidPer <> '' then
    begin
    Cle := GuidPerDos+';STE;1;'+Lafct+';'+GuidPer ;
    if GetField('ANL_TYPEDOS') = 'DP' then
      // AGLLanceFiche('DP','FICHINTERVENTION', Cle, Cle, 'ACTION=CREATION;'+Cle+';'+FormJur)
      AGLLanceFiche('DP','FICHINTERVENTION', Cle, '', 'ACTION=CREATION;'+Cle+';'+FormJur)
    // lien de type juridique (associés, administrateurs...)
    else
       AGLLanceFiche('YY','FICHELIEN', Cle, Cle,'CREATION;'+Cle+';'+FormJur);
    // limite l'affichage à la fonction qu'on vient d'alimenter
    THValComboBox(GetControl('ANL_FONCTION')).value := Lafct;
    // actualisation
    AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
    end;
end;


procedure TOF_ANNULIEN.BTLIENANNU_OnClick(Sender : TObject);
// affectée uniquement quand Fct <> 'JUR' et <> 'DOS'
// dans ce cas, pas de popup menu dans la barre d'outils
// (le clic du bouton appelle directement le lien)
var Cle, Cle2, GuidPer, Msg, Lafonc : String;
begin
  // Permet par exple de créer un CCS lorsqu'on a CSS et CCT dans la combo,
  // et que la fiche a été ouverte avec CCT en paramètre (voir BTCAC dans Juridique)
  // Rq : seul LIENPERSONNE n'a pas la zone ANL_FONCTION, mais n'accède pas à cette proc
  Lafonc := THValComboBox(GetControl('ANL_FONCTION')).value;
  if Lafonc='' then Lafonc := Fct;

  // Traitement d'un lien juridique particulier (forcer existence d'un actionnaire)
  // #### idiot puisqu'on peut créer le lien ADM dans le popup, donc sans passer
  // par là !!
  if Lafonc='ADM' then
    // Administrateur
    BEGIN
      if Ecran.Name = 'LIENDETJURIDIQUE' then
          GuidPer := AglLanceFiche('DP','LIENANNUAIRE',
               'GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafonc, '', '')
      else
          GuidPer := AglLanceFiche('DP','LIENANNUAIRE',
               'GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafonc, '', Lafonc);

    if (GuidPer <> '') then
      begin
      Cle  := GuidPerDos+';STE;1;'+ Lafonc +';'+GuidPer;
      Cle2 := GuidPerDos+';STE;1;'+'ACT'+';'+GuidPer;
      // #### est-ce toujours ACT ?
      if Not ExisteSQL('SELECT ANL_GUIDPERDOS FROM ANNULIEN WHERE ANL_GUIDPERDOS="'+GuidPerDos+
       '" AND ANL_FONCTION="ACT" AND ANL_GUIDPER="'+Guidper+'"') then
        // AglLanceFiche('DP','FICHINTERJUR', Cle, Cle, 'CREATION;'+Cle+';'+FormJur)
        AglLanceFiche('YY','FICHELIEN', Cle, Cle, 'CREATION;'+Cle+';'+FormJur)
      else
        begin
        Msg := 'Cet administrateur n''est pas actionnaire, confirmez-vous sa création comme actionnaire ?';
        if PGIAsk(Msg, TitreHalley)=mrYes then
          begin
          // AglLanceFiche('DP','FICHINTERJUR', Cle2, Cle2,'CREATION;'+Cle2+';'+FormJur);
          AglLanceFiche('YY','FICHELIEN', Cle2, Cle2,'CREATION;'+Cle2+';'+FormJur);
          // AglLanceFiche('DP','FICHINTERJUR', Cle, Cle,'CREATION;'+Cle+';'+FormJur);
          AglLanceFiche('YY','FICHELIEN', Cle, Cle,'CREATION;'+Cle+';'+FormJur);
          end;
        end;
      end;
    END
  // #### test pour éliminer les liens juridiques qui nécessitent FICHINTERJUR
  else if not ExisteSQL('select JFT_FONCTION from JUFONCTION where JFT_FONCTION="'+Lafonc
    +'" and JFT_FORME="'+FormJur+'" and JFT_TYPEDOS="STE" and JFT_TIERS<>"X"') then
  {else if (Lafonc<>'ASS') and (Lafonc<>'ACT') and (Lafonc<>'GRT')
    and (Lafonc<>'CCT') and (Lafonc<>'CCS') then }
    BEGIN
      if Ecran.Name = 'LIENDETJURIDIQUE' then
        GuidPer := AGLLanceFiche('DP','LIENANNUAIRE',
                     'GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafonc,'', '')
      else
        GuidPer := AGLLanceFiche('DP','LIENANNUAIRE',
                        'GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafonc,'', Lafonc);

    if GuidPer <> '' then
      begin
      Cle := GuidPerDos+';DP;1;'+ Lafonc +';'+GuidPer;
      // AGLLanceFiche('DP','FICHINTERVENTION', St, '0;DEF;0;DEF;0', 'CREATION;'+Cle);
      // AGLLanceFiche('DP','FICHINTERVENTION', Cle, Cle, 'ACTION=CREATION;'+Cle+';'+FormJur);
      // AGLLanceFiche('DP','FICHINTERVENTION', Cle, '', 'ACTION=CREATION;'+Cle+';'+FormJur);
      AGLLanceFiche('DP','FICHINTERVENTION', Cle, '', 'ACTION=CREATION;'+Cle+';'+FormJur+';'+GrpFiscal);
      end;
    END
  else
    // lien de type juridique
    BEGIN
      // BM FQ11607 'GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafonc,'', Lafonc);
      //if Ecran.Name = 'LIENDETJURIDIQUE' then
        GuidPer := AGLLanceFiche('DP','LIENANNUAIRE',
                        'GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafonc,'', '');
      //else
      // GuidPer := AGLLanceFiche('DP','LIENANNUAIRE',
      //                'GUIDPERDOS='+GuidPerDos+';ANLFONCTION='+Lafonc,'', LaFonc);
    if GuidPer <> '' then
      begin
      Cle := GuidPerDos+';STE;1;'+ Lafonc +';'+GuidPer;
      //AGLLanceFiche('DP','FICHINTERJUR', Cle, Cle, 'ACTION=CREATION;'+Cle+';'+FormJur);
      AGLLanceFiche('YY','FICHELIEN', Cle, Cle, 'ACTION=CREATION;'+Cle+';'+FormJur);
      end;
    END;

  // actualisation
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


procedure TOF_ANNULIEN.OnDblClickFListe(Sender : TObject);
var Cle: String;
begin
  // Vérif pour traitement de la clé
  If Not ListeCorrecte('ANL_GUIDPER', 'Code de la fiche personne') then exit;
  If Not ListeCorrecte('ANL_GUIDPERDOS', 'N° dossier général') then exit;
  if VarIsNull(GetField('ANL_GUIDPERDOS')) then exit;

  // détail du lien
  if (Ecran.Name<>'LIENINTERDOS') then // ou Fct<>'DOS', mais éliminerait aussi LIENPERSONNE
    BEGIN
    // #### serait plus simple de rendre les champs obligatoires
    //      dans les listes comme DPANNUPERS et DPANNULIEN
    If Not ListeCorrecte('ANL_FONCTION', 'Nature de lien') then exit;
    If Not ListeCorrecte('ANL_TYPEDOS', 'Type de dossier') then exit;
    //If Not ListeCorrecte('ANL_FORME', 'Forme du dossier') then exit;

    // clé et début des param lus dans OnArgument de la tom annulien
    Cle := GetField('ANL_GUIDPERDOS')
      +';'+GetField('ANL_TYPEDOS')
      +';1' // NoOrdre
      +';'+GetField('ANL_FONCTION')
      +';'+GetField('ANL_GUIDPER');

    // lien de type DP (annuaire général, fiscal, social)
    if GetField('ANL_TYPEDOS') = 'DP' then
      // ne pas supprimer Cle en paramètre "Lequel" car sinon n'ouvre pas le bon
      AGLLanceFiche('DP','FICHINTERVENTION', Cle, Cle, 'ACTION=MODIFICATION;'+Cle+';'+FormJur)
    // lien de type juridique (associés, administrateurs...)
    else
      //AGLLanceFiche('DP','FICHINTERJUR', Cle, Cle, 'ACTION=MODIFICATION;'+Cle+';'+FormJur);
      AGLLanceFiche('YY','FICHELIEN', Cle, Cle, 'ACTION=MODIFICATION;'+Cle+';'+FormJur);
    END
  // ou fiche annuaire (LIENINTERDOS appelait FICHESIGNALE=> idiot pour un org. fiscal...)
  else
    BEGIN
    AGLLanceFiche('YY','ANNUAIRE',GetField('ANL_GUIDPER'), GetField('ANL_GUIDPER'),
      'ACTION=MODIFICATION') ;
    END;

  // rafraichissement du mul
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


function TOF_ANNULIEN.ListeCorrecte(nomchp, titrechp: String): Boolean;
begin
  Result := True;
  //  if Not GetDataSet.FieldExists(nomchp) then => pas possible (du moins si pas d'enreg...)
  If Not ChampEstDansQuery(nomchp, TFMul(Ecran).Q) then
    begin
    Result := False;
    PGIInfo('La colonne "'+titrechp+'" ne figure pas dans votre paramètrage de liste.'+#13+#10
     +'Veuillez la rajouter dans "Afficher les colonnes suivantes".', TitreHalley);
    TButton(GetControl('BPARAMLISTE')).Click;
    end;
end;


procedure TOF_ANNULIEN.BTANNUAIRE_OnClick(Sender: TObject);
begin
  if VarIsNull(GetField('ANL_GUIDPER')) then exit;

  AGLLanceFiche('YY','ANNUAIRE', GetField('ANL_GUIDPER'), GetField('ANL_GUIDPER'),
    'ACTION=MODIFICATION') ;
  // dans le script, on avait FICHESIGNALE au lieu ANNUAIRE :
  //ouvrefiche('DP','FICHESIGNALE',GetChamp('ANL_GUIDPER'), GetChamp('ANL_GUIDPER'),
  // 'ACTION=MODIFICATION;DP')
end;


procedure TOF_ANNULIEN.PurgePopup;
var Fin, i : Integer;
begin
  if GetControl('POPMENUJUR')=Nil then exit;
  // ou Ecran.Name<>LIENDETJURIDIQUE, LIENJURIACTION, LIENINTERJURI

  Fin := TPopUpMenu(GetControl('POPMENUJUR')).Items.Count-1;
  // purge popup
  for i:=Fin downto 0 do
    TPopUpMenu(GetControl('POPMENUJUR')).Items.Delete(i);
end;


procedure TOF_ANNULIEN.Form_OnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of

  // Nouveau lien : Ctrl + N
  78 : if (Shift = [ssCtrl]) and GetControl('BTLIENANNU').Visible then
         begin
         TToolBarButton97(GetControl('BTLIENANNU')).Click;
         // Rq : le bouton a parfois un popup, on pourrait tester
         // if TToolBarButton97(GetControl('BTLIENANNU')).DropDownMenu<>Nil then ...
         end;

  // Suppression : Ctrl + Suppr
  VK_DELETE :
       if (Shift = [ssCtrl]) and GetControl('BDELETE').Visible then
         BDELETE_OnClick(Nil);

  else
       TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
  end;
end;

Initialization
registerclasses([TOF_ANNULIEN]) ;
end.
