{***********UNITE************************************************************************************
Modifié le ... : 20/11/2000 par JCF pour gestion zones libres
Modifié le ... : 22/11/2000 par AV pour implémentation champ ET_SURSITE
Modifié le ... : 17/07/2002 par CT pour implémentation de la gestion multi-dépôt par établissement
Modifié le ... : 04/06/2003 par NA pour unicité du libellé
Modifié le ... : 04/09/2007 par FC pour ajouter la duplication d'un établissement
*****************************************************************************************************}
unit UTomEtabliss;

interface

uses Classes, UTOM, Ent1, Forms, Hctrls, HEnt1, UtilPGI, Extctrls,
  UTob, StdCtrls, SysUtils, HMsgBox, controls, Math,
  {$IFDEF EAGLCLIENT}
  eFiche, eFichList, MaineAgl,UtileAGL,
  {$ELSE}
  db,DBCtrls,   {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} Fiche, Fe_Main, FichList,
  {$ENDIF}
  {$IFDEF GIGI}
  Dicobtp, dpTofDossierSel, dpJurOutils, // $$$JP 20/11/04
  {$ENDIF}
  {$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
  EntGC, UtilGC,
  {$ENDIF PGIMAJVER}
  {$ENDIF}
  M3FP, ParamSoc,
  eventDecla //LM20070201
  ;

type
  Tom_Etabliss = class(TOM)
  private
    Etablis_sursite: boolean;
    Error_field,ChangeLibelle: boolean;

    NTP_ : Tob ; //+LM20070201
    evt : TEventDecla ;
    InfoTP : Boolean ; //-LM20070201
    Dupli : Boolean; //FC20070904

    //function IsMouvementer: Boolean;//RM2 Remplacement par la fonction IsEtabMouvementer de type public
                                      //Il faut passer en paramètre le numéro de l'etablissement et la
                                      //fonction renvoi si il est mouvementer ou non.
    {$IFDEF GCGC}
    {$IFNDEF PGIMAJVER}
    procedure AffichePhoto;
    {$ENDIF PGIMAJVER}
    {$ENDIF}

    // $$$ JP 06/04/2004 - sélection dossier cabinet
//    {$IFNDEF EAGLCLIENT}
    {$IFDEF GIGI}
    procedure DossierElipsisClick (Sender:TObject);
    procedure DossierDelClick (Sender:TObject);
    {$ENDIF}
//    {$ENDIF}
    procedure TPOnEnter (Sender:TObject);//LM20070201
    procedure TPOnExit (Sender:TObject); //LM20070201
    {$IFDEF PAIEGRH}
    procedure DupliquerEtab(Sender: TObject); //FC20070904
    {$ENDIF}
  public
    TobListe, TobUtilisable: TOB;
    GLISTE, GUTILISABLE: THGrid;
    procedure OnDeleteRecord; override;
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure ClickFlecheDroite;
    procedure ClickFlecheGauche;
    procedure ClickFlecheTous;
    procedure ClickFlecheAucun;
    procedure ClickFlecheHaut;
    procedure ClickFlecheBas;
    procedure RefreshGrid(posListe, posUtilisable: integer);
    procedure RefreshBouton;
    procedure DisableBoutons;
    procedure ClickInsertDepot(Depot: string);
    procedure ValideListeDepot;
  end;

  procedure DeleteEtablissement(CodeEtab: String);//RM2 20070702
  function IsEtabMouvementer(CodeEtab: String; var LastErrorMsg: HString): integer;//RM2 cf IsMouvementer

const
  // libellés des messages
  TexteMessage: array[1..19] of string = (
    {1}'Vous ne pouvez pas supprimer l''établissement par défaut.',
    {2}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des écritures comptables',
    {3}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des écritures analytiques',
    {4}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des écritures budgétaires',
    {5}'Ce code existe déjà ! Vous devez le mettre à jour',
    {6}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des pièces',
    {7}'Vous avez atteint le nombre maximum d''établissements gérés sur site pour votre licence :',
    {8}'Vous ne pouvez pas supprimer cet établissement, il est référencé par le logiciel de paie',
    {9}'Il existe déjà un établissement avec le même libellé',
    {10}'10',
    {11}'11',
    {12}'12',
    {13}'13',
    {14}'14',
    {15}'15',
    {16}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des affaires',
    {17}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des ressources',
    {18}'Vous ne pouvez pas supprimer cet établissement, il est référencé dans la table afcumul',
    {19}'Vous ne pouvez pas supprimer cet établissement, il est référencé dans les immobilisations'
    );
  // les différents boutons de navigation pour les dépôts liés
  BTN_DROIT = 'DROITE';
  BTN_GAUCHE = 'GAUCHE';
  BTN_HAUT = 'HAUT';
  BTN_BAS = 'BAS';
  BTN_TOUS = 'TOUS';
  BTN_AUCUN = 'AUCUN';
  BTN_INSERT = 'INSERT';
  GRD_LISTE = 'GLISTE';
  GRD_UTILISABLE = 'GUTILISABLE';

implementation

// $$$JP 02/06/04: on a besoin du THDBEdit
//{$IFNDEF EAGLCLIENT}
uses hdb, htb97, BtpUtil,HrichOle
{$IFDEF PAIEGRH}
,pgOutils
{$ENDIF}
;
//{$ENDIF}


procedure TOM_Etabliss.OnArgument(Arguments: string);
{var Critere,ChampMul,ValMul,DepotsLies,listeDepotsLies : string;
    x,i,iListe,iDepot : integer;
    QQ : TQuery ;
    TobDepot : TOB ;
    bTrouve : boolean ; }
//{$IFNDEF EAGLCLIENT}

  var i:integer ; //LM20070201
{$IFDEF GIGI}
  var     NoDossierEdit :THDBEdit;
{$ENDIF}
//{$ENDIF}
begin
  inherited;
  AppliqueFontDefaut (THRichEditOle(GetControl('ET_BLOCNOTE')));
  Dupli := False; //FC20070904
  //+LM20070201
  evt := TEventDecla.create (Ecran) ;

  NTP_:=tob.create ('NETABLISSEMENT', nil, -1) ;
  if pos('INFOTP', upperCase(Arguments))>0 then
  begin
    InfoTP:=True ;
    SetControlProperty('PEBTTP','TabVisible', True);

    for i:=1 To NTP_.NbChamps do //contrôle modif de zone
    begin
      evt.Rebranche(NTP_.GetNomChamp(i), 'OnEnter', TPOnEnter);
      evt.Rebranche(NTP_.GetNomChamp(i), 'OnExit', TPOnExit);
    end ;
  end ;
  //-LM20070201

  {récup des arguments }
  {    Inutile en théorie
  Repeat
      Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
      if Critere<>'' then
          begin
          x:=pos('=',Critere);
          if x<>0 then
             begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='ACTION' then
                begin
                if ValMul = 'MODIFICATION' then FAction := TaModif else FAction := TaCreat ;
               end;
             end;
          end;
  until  Critere='';
  }
  {$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}

  GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'ET_LIBREET', 10, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'ET_VALLIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'ET_DATELIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'ET_BOOLLIBRE', 3, '');
  // mcd 27/11/01 GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'ET_TEXTELIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'ET_CHARLIBRE', 3, '');

  SetActiveTabSheet('PGeneral');
  {$ENDIF PGIMAJVER}
  {$ENDIF}
  with TFFicheListe(Ecran) do HelpContext := 110000216;
  if (Ecran <> nil) and (Ecran is TFFiche) and
    (TFFiche(Ecran).TypeAction = taConsult) then
  begin
    DisableBoutons;
  end;

  // $$$JP 06/04/04 - dossier de production associé à l'établissement: uniquement dans base commune en multi-dossier
  // $$$ A FAIRE EAGL
//{$IFNDEF EAGLCLIENT}
  if (V_PGI.DefaultSectionName <> '') and (V_PGI.InBaseCommune = FALSE) then
  begin
//{$ENDIF}
     SetControlVisible ('PDOSSIER', FALSE);
//{$IFNDEF EAGLCLIENT}
  end
  else
  begin
       {$IFDEF GIGI}
       NoDossierEdit := THDBEdit (GetControl ('ET_NODOSSIER'));
       if NoDossierEdit <> nil then
       begin
            NoDossierEdit.OnElipsisClick := DossierElipsisClick;
            TToolBarButton97 (GetControl ('BDELNODOSSIER')).OnClick := DossierDelClick;
       end;
       {$ELSE}
          SetControlVisible ('PDOSSIER', FALSE);
       {$ENDIF}
  end;
//{$ENDIF}
  //FC20070904
{$IFDEF PAIEGRH}
  SetControlVisible('BDUPLIQUER',True);
  TToolBarButton97 (GetControl ('BDUPLIQUER')).OnClick := DupliquerEtab;
{$ENDIF}
end;

// $$$JP 06/04/04 - sélection dossier cabinet
//{$IFNDEF EAGLCLIENT}
{$IFDEF GIGI}
procedure Tom_Etabliss.DossierElipsisClick (Sender:TObject);
var
   sRetour  :string;
begin
     // retourne NoDossier;CodePer;Nom1
     // $$$ JP 07/06/04 - ne plus filtrer les dossier cabinet: "cabinet" ne signifie pas entité juridique, mais bien base du cabinet
     sRetour := DP_SelectUnDossier (GetControlText ('ET_NODOSSIER'), FALSE);
     if sRetour <> '' then
     begin
          ModeEdition (DS);
          SetField ('ET_NODOSSIER', ReadTokenSt (sRetour));
     end;
end;

procedure Tom_Etabliss.DossierDelClick (Sender:TObject);
begin
     if GetField ('ET_NODOSSIER') <> '' then
        if PgiAsk ('Les traitements nécessitant la connaissance du dossier de production ne seront plus actifs' + #10 + ' Etes-vous sûr de ne pas associer de dossier à cet établissement?', 'Lien établissement/dossier') = mrYes then
        begin
             ModeEdition (DS);
             SetField ('ET_NODOSSIER', '');
        end;
end;
{$ENDIF}
//{$ENDIF}

procedure Tom_Etabliss.OnClose;
begin
  TobListe.Free;
  TobUtilisable.Free;
  NTP_.Free ;//LM20070201
  evt.free;//LM20070201
  inherited;

end;

procedure Tom_Etabliss.OnDeleteRecord;
var
  Etab: string;
begin
  Etab := GetField('ET_ETABLISSEMENT');
  if Etab = '' then Exit;
  if Etab = VH^.EtablisDefaut then
  begin
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;
  LastError := IsEtabMouvementer(Etab, LastErrorMsg);
  if LastError > 0 then
    exit;
  DeleteEtablissement(Etab);
  inherited;
end;

function IsEtabMouvementer(CodeEtab: String; var LastErrorMsg: HString): integer;
var Sql: string;
begin
  Result := 0;
  Sql := 'Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="' + CodeEtab + '" And ' +
    '(Exists (Select E_ETABLISSEMENT From ECRITURE Where E_ETABLISSEMENT="' + CodeEtab + '"))';
  if ExisteSQL(SQL) then
  begin
    Result := 2;
    LastErrorMsg := TexteMessage[Result];
    Exit;
  end;

  Sql := 'Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="' + CodeEtab + '" And ' +
    '(Exists (Select Y_ETABLISSEMENT From ANALYTIQ Where Y_ETABLISSEMENT="' + CodeEtab + '"))';
  if ExisteSQL(SQL) then
  begin
    Result := 3;
    LastErrorMsg := TexteMessage[Result];
    Exit;
  end;

  if not EstBasePclAllegee then  // CA - 11/10/2005 pour test base allégée PCL
  begin
    Sql := 'Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="' + CodeEtab + '" And ' +
      '(Exists (Select GP_ETABLISSEMENT From PIECE Where GP_ETABLISSEMENT="' + CodeEtab + '"))';
    if ExisteSQL(SQL) then
    begin
      Result := 6;
      LastErrorMsg := TexteMessage[Result];
      Exit;
    end;
  end;

  Sql := 'Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="' + CodeEtab + '" And ' +
    '(Exists (Select BE_ETABLISSEMENT From BUDECR Where BE_ETABLISSEMENT="' + CodeEtab + '"))';
  if ExisteSQL(SQL) then
  begin
    Result := 4;
    LastErrorMsg := TexteMessage[Result];
    Exit;
  end;

  //PAIE
  Sql := 'Select ETB_ETABLISSEMENT From ETABCOMPL Where ETB_ETABLISSEMENT="' + CodeEtab + '"';
  if ExisteSQL(SQL) then
  begin
    Result := 8;
    LastErrorMsg := TexteMessage[Result];
    Exit;
  end;

  // IMMO Fiche 19601
  if ExisteSQL ('SELECT I_ETABLISSEMENT FROM IMMO WHERE I_ETABLISSEMENT="'+ CodeEtab +'"') then
  begin
    Result := 19;
    LastErrorMsg := TexteMessage[Result];
    Exit;
  end;


  if not EstBasePclAllegee then  // CA - 11/10/2005 pour test base allégée PCL
  begin
    // mcd 16/11/2004 equalité 11695 .. ilf aut prendre en compte les tables affaire
    Sql := 'Select AFF_ETABLISSEMENT From AFFAIRE Where AFF_ETABLISSEMENT="' + CodeEtab + '"';
    if ExisteSQL(SQL) then
    begin
      Result := 16;
  {$ifdef GIGI}
      LastErrorMsg := TraduitGa(TexteMessage[Result]);
  {$else}
      LastErrorMsg := TexteMessage[Result];
  {$endif}
      Exit;
    end;
    Sql := 'Select ARS_ETABLISSEMENT From RESSOURCE Where ARS_ETABLISSEMENT="' + CodeEtab + '"';
    if ExisteSQL(SQL) then
    begin
      Result := 17;
  {$ifdef GIGI}
      LastErrorMsg := TraduitGa(TexteMessage[Result]);
  {$else}
      LastErrorMsg := TexteMessage[Result];
  {$endif}
      Exit;
    end;
    Sql := 'Select ACU_ETABLISSEMENT From AFCUMUL Where ACU_ETABLISSEMENT="' + CodeEtab + '"';
    if ExisteSQL(SQL) then
    begin
      Result := 18;
  {$ifdef GIGI}
      LastErrorMsg := TraduitGa(TexteMessage[Result]);
  {$else}
      LastErrorMsg := TexteMessage[Result];
  {$endif}
      Exit;
    end;
  end;
end;

procedure DeleteEtablissement(CodeEtab: String);
var SQL: String;
begin
  {$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
  if (not VH_GC.GCMultiDepots) then
  begin
    // Suppression du dépôt
    if ExisteSQL('SELECT GDE_DEPOT FROM DEPOTS WHERE GDE_DEPOT="' + CodeEtab + '"') then
    begin
      SQL := 'DELETE FROM DEPOTS WHERE GDE_DEPOT="' + CodeEtab + '"';
      ExecuteSQL(SQL);
    end;
  end;
  {$ENDIF PGIMAJVER}
  {$ENDIF}
  ExecuteSQL('DELETE FROM CONTACT WHERE C_TYPECONTACT="ET" AND C_AUXILIAIRE="' + CodeEtab + '"');
  ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="ET" AND LO_IDENTIFIANT="' + CodeEtab + '"');
  ExecuteSQL('delete from NETABLISSEMENT where NTP_ETABLISSEMENT="' + CodeEtab + '"'); //LM20070201
end;

{$IFDEF GCGC}
{$IFNDEF PGIMAJVER}

procedure TOM_Etabliss.AffichePhoto;
var QQ: TQuery;
  SQL: string;
  CC: TImage;
  IsJpeg: boolean;
begin
  SetControlVisible('P_PHOTO', false);
  CC := TImage(GetControl('LAPHOTO'));
  if CC = nil then exit;
  if VH_GC.GCPHOTOFICHE = '' then exit;
  SQL := 'SELECT * from LIENSOLE where LO_TABLEBLOB="ET" AND LO_IDENTIFIANT="' + getField('ET_ETABLISSEMENT') + '"';
  SQL := SQL + ' AND (LO_QUALIFIANTBLOB="PHO" OR  LO_QUALIFIANTBLOB="PHJ" OR  LO_QUALIFIANTBLOB="VIJ") AND LO_EMPLOIBLOB="' + VH_GC.GCPHOTOFICHE + '"';
  QQ := OpenSQL(SQL, true);
  if not QQ.EOF then
  begin
    IsJpeg := ((QQ.Findfield('LO_QUALIFIANTBLOB').asString = 'PHJ') or (QQ.Findfield('LO_QUALIFIANTBLOB').asString = 'VIJ'));
    LoadBitMapFromChamp(QQ, 'LO_OBJET', CC, IsJpeg);
    if CC.Picture <> nil then SetControlVisible('P_PHOTO', true);
  end;
  ferme(QQ);
end;
{$ENDIF PGIMAJVER}
{$ENDIF}

procedure Tom_Etabliss.OnNewRecord;
begin
  inherited;
  SetField('ET_SURSITE', 'X');
  Etablis_sursite := False;

  NTP_.InitValeurs;//LM20070201
  NTP_.PutEcran(Ecran); //LM20070201

  // $$$ JP 07/04/04 - par défaut, pas de dossier attaché à l'établissement
  {$IFDEF GIGI}
  SetField ('ET_NODOSSIER', '');
  {$ENDIF}


end;

procedure Tom_Etabliss.OnLoadRecord;
{$IFDEF GCGC}
{$IFNDEF PGIMAJVER}
var DepotsLies, listeDepotsLies: string;
  iListe: integer;
  QQ: TQuery;
  bTrouve: boolean;
  {$ENDIF PGIMAJVER}
{$ENDIF}
begin
  inherited;
  if ((EstSerie(S3)) or (EstSerie(S5))) then
  begin
    SetControlVisible('ET_LANGUE', False);
    SetControlVisible('TET_LANGUE', False);
  end;
  // ****************************CT le 17/07/2002 **********************
  // Modif pour pouvoir lier des dépôts à l'établissement
  // Chargement de la liste des établissements et du paramétrage existant.
  {$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
  if (Ecran.Name <> 'GCETABLISS') and (Ecran.Name <> 'YYETABLISS') and (Ecran.Name <> 'ETABLISSEMENT')
   and (Ecran.Name <> 'NETABLISSEMENT') then //LM20070201
  begin
    if VH_GC.GCMultiDepots then // Nous sommes en multi-dépôts
    begin
      SetControlVisible('ONG_LIAISONDEPOT', True);
      TobListe := TOB.CREATE('Liste dépôts', nil, -1);
      TobUtilisable := TOB.CREATE('Dépôts utilisables', nil, -1);
      QQ := OpenSQL('select GDE_DEPOT,GDE_LIBELLE,GDE_ABREGE from DEPOTS order by GDE_DEPOT', True);
      // Chargement paramétrage sauvegardé
      ListeDepotsLies := GetControlText('ET_DEPOTLIE');

      if ListeDepotsLies <> '' then // (pas fonctionnement de GESCOM) l'établissement est lié à certain dépôts, pas tous
      begin
        // Chargement de la liste des établissements dans TobListe
        // et bascule des dépots se trouvant dans ET_DEPOTLIE dans TobUtilisable trié suivant l'ordre définit dans ET_DEPOTLIE
        TobListe.LoadDetailDB('DEPOTS', '', '', QQ, False);
        DepotsLies := ReadTokenSt(ListeDepotslies);
        while DepotsLies <> '' do
        begin
          iListe := 0;
          bTrouve := False;
          while (not bTrouve) and (iListe < TobListe.Detail.Count) do
          begin
            bTrouve := (TobListe.Detail[iListe].GetValue('GDE_DEPOT') = DepotsLies);
            inc(iListe);
          end;
          if bTrouve then TobListe.Detail[iListe - 1].ChangeParent(TobUtilisable, -1);
          DepotsLies := ReadTokenSt(ListeDepotslies);
        end;
      end
      else // (fonctionnement de GESCOM) établissement lié à tous les dépôts
      begin
        // tous les dépôts sont utilisables trié par le code dépôt
        TobUtilisable.LoadDetailDB('DEPOTS', '', '', QQ, False);
      end;
      Ferme(QQ);
      // Affichage des tobs
      GLISTE := THGrid(GetControl('GLISTE'));
      GUTILISABLE := THGrid(GetControl('GUTILISABLE'));
      TobUtilisable.PutGridDetail(GUTILISABLE, False, False, 'GDE_DEPOT;GDE_LIBELLE', True);
      TobListe.PutGridDetail(GLISTE, False, False, 'GDE_DEPOT;GDE_LIBELLE', True);
      GLISTE.ColWidths[0] := 30;
      GLISTE.ColAligns[0] := taCenter;
      GLISTE.ColWidths[1] := 208;
      GLISTE.ColAligns[1] := taLeftJustify;
      GUTILISABLE.ColWidths[0] := 30;
      GUTILISABLE.ColAligns[0] := taCenter;
      GUTILISABLE.ColWidths[1] := 208;
      GUTILISABLE.ColAligns[1] := taLeftJustify;
    end
      // Nous sommes en mono-dépôt, l'onglet de liaison des dépôts n'est pas visible
    else SetControlVisible('ONG_LIAISONDEPOT', False);
  end;
  {$ELSE}
     SetControlVisible('ONG_LIAISONDEPOT', False);
  {$ENDIF PGIMAJVER}

  {$ELSE}
  SetControlVisible('ONG_LIAISONDEPOT', False);
  {$ENDIF}

  // affiche Photo
  {$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
  if VH_GC.GCPHOTOFICHE <> '' then AffichePhoto else
  {$ENDIF PGIMAJVER}
    {$ENDIF}
    SetControlVisible('P_PHOTO', false);

  if InfoTP then //LM20070201
  begin
    NTP_.p('NTP_ETABLISSEMENT', GetField('ET_ETABLISSEMENT')) ;
    if not NTP_.LoadDB then NTP_.InitValeurs;
    NTP_.PutEcran(Ecran);
  end ;

end;

procedure TOM_Etabliss.OnChangeField(F: TField);
var QQ: TQuery;
  SQL: string;
  {$IFDEF GIGI}
  TOBLibDos    :TOB; // $$$ JP 07/04/04 - libellé dossier
  {$ENDIF}
begin
  inherited;
  Error_field := False;
  // 1 - En création, on vérifie que ce code n'existe pas
  if (F.FieldName = 'ET_ETABLISSEMENT') and (DS.State in [dsInsert]) and (not Dupli) then   //FC20070904
  begin
    SQL := 'SELECT ET_ETABLISSEMENT from ETABLISS where ET_ETABLISSEMENT="' + getField('ET_ETABLISSEMENT') + '"';
    QQ := OpenSQL(SQL, true);
    if not QQ.EOF then
    begin
      LastError := 5;
      LastErrorMsg := TexteMessage[LastError];
      Error_field := True;
      SetFocusControl('ET_ETABLISSEMENT');
    end;
    ferme(QQ);
  end;
  // Le libellé de l'établissement doit être unique
  if F.FieldName = 'ET_LIBELLE' then
  begin
    SQL := 'SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_LIBELLE="' + GetField('ET_LIBELLE') + '"'
      + ' AND ET_ETABLISSEMENT<>"' + GetField('ET_ETABLISSEMENT') + '"';
    if ExisteSQL(SQL) then
    begin
      LastError := 9;
      LastErrorMsg := TexteMessage[LastError];
      Error_field := True;
      SetFocusControl('ET_LIBELLE');
    end;
  end;
  if (F.FieldName='ET_LIBELLE') AND (Ds.State in [DsEdit]) then //SB 17/09/2003
    ChangeLibelle:=True;

  // $$$ JP 07/04/04 - affichage nom dossier
  {$IFDEF GIGI}
  if (F.FieldName = 'ET_NODOSSIER') then
  begin
       if DS.State in [DsEdit] then
          PgiInfo ('Le lien de l''établissement vers le dossier est déterminant' + #10 + ' En particulier, le transfert comptable aux entités juridiques du cabinet devra être ré-initialisé', 'Lien établissement - dossier');
       TOBLibDos := TOB.Create ('nom dossier', nil, -1);
       try
          TOBLibDos.LoadDetailFromSQL ('SELECT DOS_LIBELLE FROM DOSSIER WHERE DOS_NODOSSIER="' + GetField ('ET_NODOSSIER') + '"');
          if TOBLibDos.Detail.Count > 0 then
              SetControlText ('LNOMDOSSIER', TOBLibDos.Detail [0].GetValue ('DOS_LIBELLE'))
          else
              SetControlText ('LNOMDOSSIER', '');
       finally
              TOBLibDos.Free;
       end;
  end;
  {$ENDIF}
end;

procedure TOM_Etabliss.OnUpdateRecord;
{$IFDEF GCGC}
{$IFNDEF PGIMAJVER}
var TOB_DEPOT: TOB;
  QQ: TQuery;
  Nb: integer;
  {$ENDIF PGIMAJVER}
  {$ENDIF}
begin
  inherited;
  SetField('ET_ETABLISSEMENT', Trim(GetField('ET_ETABLISSEMENT')));

  if InfoTP then  //LM20070201
  begin
    NTP_.p('NTP_ETABLISSEMENT', GetField('ET_ETABLISSEMENT')) ;
    NTP_.GetEcran(Ecran);
    NTP_.InsertOrUpdateDB ;
  end ;

  {$IFDEF COMPTA}
// ajout me pour la sychronisation
  MAJHistoparam ('ETABLISS', Trim(GetField('ET_ETABLISSEMENT')));
  {$ENDIF}

  {$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
// DBR Pour supprimer les conseils et avertissements
  Tob_Depot := nil;

  if (not VH_GC.GCMultiDepots) and (Error_field = False) then // Si Mono-Dépôt et pas d'erreur
  begin
    // En mono-dépôt, le dépôt doit être identique à l'établissement, donc on l'écrase dans tous les cas
    TOB_DEPOT := TOB.Create('DEPOTS', nil, -1); // Création de la TOB des dépôts
    TOB_DEPOT.PutValue('GDE_DEPOT', GetField('ET_ETABLISSEMENT'));
    TOB_DEPOT.PutValue('GDE_LIBELLE', GetField('ET_LIBELLE'));
    TOB_DEPOT.PutValue('GDE_ABREGE', GetField('ET_ABREGE'));
    TOB_DEPOT.PutValue('GDE_ADRESSE1', GetField('ET_ADRESSE1'));
    TOB_DEPOT.PutValue('GDE_ADRESSE2', GetField('ET_ADRESSE2'));
    TOB_DEPOT.PutValue('GDE_ADRESSE3', GetField('ET_ADRESSE3'));
    TOB_DEPOT.PutValue('GDE_CODEPOSTAL', GetField('ET_CODEPOSTAL'));
    TOB_DEPOT.PutValue('GDE_VILLE', GetField('ET_VILLE'));
    TOB_DEPOT.PutValue('GDE_PAYS', GetField('ET_PAYS'));
    TOB_DEPOT.PutValue('GDE_TELEPHONE', GetField('ET_TELEPHONE'));
    TOB_DEPOT.PutValue('GDE_FAX', GetField('ET_FAX'));
    TOB_DEPOT.PutValue('GDE_EMAIL', GetField('ET_EMAIL'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP1', GetField('ET_LIBREET1'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP2', GetField('ET_LIBREET2'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP3', GetField('ET_LIBREET3'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP4', GetField('ET_LIBREET4'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP5', GetField('ET_LIBREET5'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP6', GetField('ET_LIBREET6'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP7', GetField('ET_LIBREET7'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP8', GetField('ET_LIBREET8'));
    TOB_DEPOT.PutValue('GDE_LIBREDEP9', GetField('ET_LIBREET9'));
    TOB_DEPOT.PutValue('GDE_LIBREDEPA', GetField('ET_LIBREETA'));
    TOB_DEPOT.PutValue('GDE_DATELIBRE1', GetField('ET_DATELIBRE1'));
    TOB_DEPOT.PutValue('GDE_DATELIBRE2', GetField('ET_DATELIBRE2'));
    TOB_DEPOT.PutValue('GDE_DATELIBRE3', GetField('ET_DATELIBRE3'));
    TOB_DEPOT.PutValue('GDE_BOOLLIBRE1', GetField('ET_BOOLLIBRE1'));
    TOB_DEPOT.PutValue('GDE_BOOLLIBRE2', GetField('ET_BOOLLIBRE2'));
    TOB_DEPOT.PutValue('GDE_BOOLLIBRE3', GetField('ET_BOOLLIBRE3'));
    TOB_DEPOT.PutValue('GDE_CHARLIBRE1', GetField('ET_CHARLIBRE1'));
    TOB_DEPOT.PutValue('GDE_CHARLIBRE2', GetField('ET_CHARLIBRE2'));
    TOB_DEPOT.PutValue('GDE_CHARLIBRE3', GetField('ET_CHARLIBRE3'));
    TOB_DEPOT.PutValue('GDE_VALLIBRE1', GetField('ET_VALLIBRE1'));
    TOB_DEPOT.PutValue('GDE_VALLIBRE2', GetField('ET_VALLIBRE2'));
    TOB_DEPOT.PutValue('GDE_VALLIBRE3', GetField('ET_VALLIBRE3'));
    TOB_DEPOT.PutValue('GDE_UTILISATEUR', GetField('ET_UTILISATEUR'));
    if (DS.State = dsInsert) then TOB_DEPOT.PutValue('GDE_DATECREATION', GetField('ET_DATECREATION'));
    TOB_DEPOT.PutValue('GDE_DATEMODIF', GetField('ET_DATEMODIF'));
    TOB_DEPOT.PutValue('GDE_SURSITE', GetField('ET_SURSITE'));
    TOB_DEPOT.PutValue('GDE_SURSITEDISTANT', GetField('ET_SURSITEDISTANT'));
    TOB_DEPOT.SetAllModifie(True);
    TOB_DEPOT.InsertOrUpdateDB(FALSE);
  end
  else // sinon multi-dépôts
  begin
    if (GetParamSoc('SO_GCLIAISONAUTODEP_ETAB')) and (Error_field = False) then
    begin
      // Si nous inserons un nouveau établissement ainsi qu'un dépôt,
      // nous mettons se dépôt dans la liste des utilisables
      if (DS.State = dsInsert) then
      begin
        if GetField('ET_DEPOTLIE') <> '' then
          SetField('ET_DEPOTLIE', GetField('ET_DEPOTLIE') + ';' + GetField('ET_ETABLISSEMENT'));
      end;
      // Vérification que le dépôt correspondant à l'établissement existe
      // Si ce n'est pas le cas, on le crée
      TOB_DEPOT := TOB.Create('DEPOTS', nil, -1); // Création de la TOB des dépôts
      if not ExisteSQL('SELECT GDE_DEPOT FROM DEPOTS WHERE GDE_DEPOT="' + GetField('ET_ETABLISSEMENT') + '"') then
      begin
        TOB_DEPOT.PutValue('GDE_DEPOT', GetField('ET_ETABLISSEMENT'));
        TOB_DEPOT.PutValue('GDE_LIBELLE', GetField('ET_LIBELLE'));
        TOB_DEPOT.PutValue('GDE_ABREGE', GetField('ET_ABREGE'));
        TOB_DEPOT.PutValue('GDE_ADRESSE1', GetField('ET_ADRESSE1'));
        TOB_DEPOT.PutValue('GDE_ADRESSE2', GetField('ET_ADRESSE2'));
        TOB_DEPOT.PutValue('GDE_ADRESSE3', GetField('ET_ADRESSE3'));
        TOB_DEPOT.PutValue('GDE_CODEPOSTAL', GetField('ET_CODEPOSTAL'));
        TOB_DEPOT.PutValue('GDE_VILLE', GetField('ET_VILLE'));
        TOB_DEPOT.PutValue('GDE_PAYS', GetField('ET_PAYS'));
        TOB_DEPOT.PutValue('GDE_TELEPHONE', GetField('ET_TELEPHONE'));
        TOB_DEPOT.PutValue('GDE_FAX', GetField('ET_FAX'));
        TOB_DEPOT.PutValue('GDE_EMAIL', GetField('ET_EMAIL'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP1', GetField('ET_LIBREET1'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP2', GetField('ET_LIBREET2'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP3', GetField('ET_LIBREET3'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP4', GetField('ET_LIBREET4'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP5', GetField('ET_LIBREET5'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP6', GetField('ET_LIBREET6'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP7', GetField('ET_LIBREET7'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP8', GetField('ET_LIBREET8'));
        TOB_DEPOT.PutValue('GDE_LIBREDEP9', GetField('ET_LIBREET9'));
        TOB_DEPOT.PutValue('GDE_LIBREDEPA', GetField('ET_LIBREETA'));
        TOB_DEPOT.PutValue('GDE_DATELIBRE1', GetField('ET_DATELIBRE1'));
        TOB_DEPOT.PutValue('GDE_DATELIBRE2', GetField('ET_DATELIBRE2'));
        TOB_DEPOT.PutValue('GDE_DATELIBRE3', GetField('ET_DATELIBRE3'));
        TOB_DEPOT.PutValue('GDE_BOOLLIBRE1', GetField('ET_BOOLLIBRE1'));
        TOB_DEPOT.PutValue('GDE_BOOLLIBRE2', GetField('ET_BOOLLIBRE2'));
        TOB_DEPOT.PutValue('GDE_BOOLLIBRE3', GetField('ET_BOOLLIBRE3'));
        TOB_DEPOT.PutValue('GDE_CHARLIBRE1', GetField('ET_CHARLIBRE1'));
        TOB_DEPOT.PutValue('GDE_CHARLIBRE2', GetField('ET_CHARLIBRE2'));
        TOB_DEPOT.PutValue('GDE_CHARLIBRE3', GetField('ET_CHARLIBRE3'));
        TOB_DEPOT.PutValue('GDE_VALLIBRE1', GetField('ET_VALLIBRE1'));
        TOB_DEPOT.PutValue('GDE_VALLIBRE2', GetField('ET_VALLIBRE2'));
        TOB_DEPOT.PutValue('GDE_VALLIBRE3', GetField('ET_VALLIBRE3'));
        TOB_DEPOT.PutValue('GDE_UTILISATEUR', GetField('ET_UTILISATEUR'));
        TOB_DEPOT.PutValue('GDE_DATECREATION', GetField('ET_DATECREATION'));
        TOB_DEPOT.PutValue('GDE_DATEMODIF', GetField('ET_DATEMODIF'));
        TOB_DEPOT.PutValue('GDE_SURSITE', GetField('ET_SURSITE'));
        TOB_DEPOT.PutValue('GDE_SURSITEDISTANT', GetField('ET_SURSITEDISTANT'));
        TOB_DEPOT.InsertDB(nil, FALSE); //InsertOrUpdateDB (FALSE);
      end
      else // Si le dépôt existe déjà on demande à l'utilisateur si il veut l'écraser
      begin
        QQ := OpenSQL('SELECT Count(*) FROM DEPOTS WHERE GDE_DEPOT="' + GetField('ET_ETABLISSEMENT') + '"', True);
        Nb := QQ.Fields[0].AsInteger;
        Ferme(QQ);
        if (Nb = 1) and (DS.State = dsInsert) and (TOB_DEPOT.SelectDB('"' + GetField('ET_ETABLISSEMENT') + '"', nil) = TRUE) then
          if HShowMessage('1;;Ce dépôt existe déjà, voulez-vous l''écraser?;W;YN;N;N;;;', '', '') = mrYes then
          begin
            TOB_DEPOT.LoadDB(FALSE);
            TOB_DEPOT.PutValue('GDE_DEPOT', GetField('ET_ETABLISSEMENT'));
            TOB_DEPOT.PutValue('GDE_LIBELLE', GetField('ET_LIBELLE'));
            TOB_DEPOT.PutValue('GDE_ABREGE', GetField('ET_ABREGE'));
            TOB_DEPOT.PutValue('GDE_ADRESSE1', GetField('ET_ADRESSE1'));
            TOB_DEPOT.PutValue('GDE_ADRESSE2', GetField('ET_ADRESSE2'));
            TOB_DEPOT.PutValue('GDE_ADRESSE3', GetField('ET_ADRESSE3'));
            TOB_DEPOT.PutValue('GDE_CODEPOSTAL', GetField('ET_CODEPOSTAL'));
            TOB_DEPOT.PutValue('GDE_VILLE', GetField('ET_VILLE'));
            TOB_DEPOT.PutValue('GDE_PAYS', GetField('ET_PAYS'));
            TOB_DEPOT.PutValue('GDE_TELEPHONE', GetField('ET_TELEPHONE'));
            TOB_DEPOT.PutValue('GDE_FAX', GetField('ET_FAX'));
            TOB_DEPOT.PutValue('GDE_EMAIL', GetField('ET_EMAIL'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP1', GetField('ET_LIBREET1'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP2', GetField('ET_LIBREET2'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP3', GetField('ET_LIBREET3'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP4', GetField('ET_LIBREET4'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP5', GetField('ET_LIBREET5'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP6', GetField('ET_LIBREET6'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP7', GetField('ET_LIBREET7'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP8', GetField('ET_LIBREET8'));
            TOB_DEPOT.PutValue('GDE_LIBREDEP9', GetField('ET_LIBREET9'));
            TOB_DEPOT.PutValue('GDE_LIBREDEPA', GetField('ET_LIBREETA'));
            TOB_DEPOT.PutValue('GDE_DATELIBRE1', GetField('ET_DATELIBRE1'));
            TOB_DEPOT.PutValue('GDE_DATELIBRE2', GetField('ET_DATELIBRE2'));
            TOB_DEPOT.PutValue('GDE_DATELIBRE3', GetField('ET_DATELIBRE3'));
            TOB_DEPOT.PutValue('GDE_BOOLLIBRE1', GetField('ET_BOOLLIBRE1'));
            TOB_DEPOT.PutValue('GDE_BOOLLIBRE2', GetField('ET_BOOLLIBRE2'));
            TOB_DEPOT.PutValue('GDE_BOOLLIBRE3', GetField('ET_BOOLLIBRE3'));
            TOB_DEPOT.PutValue('GDE_CHARLIBRE1', GetField('ET_CHARLIBRE1'));
            TOB_DEPOT.PutValue('GDE_CHARLIBRE2', GetField('ET_CHARLIBRE2'));
            TOB_DEPOT.PutValue('GDE_CHARLIBRE3', GetField('ET_CHARLIBRE3'));
            TOB_DEPOT.PutValue('GDE_VALLIBRE1', GetField('ET_VALLIBRE1'));
            TOB_DEPOT.PutValue('GDE_VALLIBRE2', GetField('ET_VALLIBRE2'));
            TOB_DEPOT.PutValue('GDE_VALLIBRE3', GetField('ET_VALLIBRE3'));
            TOB_DEPOT.PutValue('GDE_UTILISATEUR', GetField('ET_UTILISATEUR'));
            TOB_DEPOT.PutValue('GDE_DATEMODIF', GetField('ET_DATEMODIF'));
            TOB_DEPOT.PutValue('GDE_SURSITE', GetField('ET_SURSITE'));
            TOB_DEPOT.PutValue('GDE_SURSITEDISTANT', GetField('ET_SURSITEDISTANT'));
            TOB_DEPOT.UpdateDB(FALSE) //InsertOrUpdateDB (FALSE);
          end;
      end;
    end;
  end;
  // DBR Pour supprimer les conseils et avertissements
  if Tob_Depot <> nil then
  TOB_DEPOT.free;
  {$ENDIF PGIMAJVER}
  {$ENDIF}
  Etablis_sursite := (GetField('ET_SURSITE') = 'X');
  If ChangeLibelle then //SB 17/09/2003
  Begin
    ExecuteSql('UPDATE ETABCOMPL SET ETB_LIBELLE="'+GetField('ET_LIBELLE')+'" '+
               'WHERE ETB_ETABLISSEMENT="'+GetField('ET_ETABLISSEMENT')+'"');
    ChangeLibelle:=False;
  End;
end;

{
if ((not VH_GC.GCMultiDepots) or ( VH_GC.GCMultiDepots and GetParamSoc('SO_GCLIAISONAUTODEP_ETAB'))) and (Error_field=False) then
   BEGIN
   // Si nous inserons un nouveau etablissement ainsi qu'un dépôt,
   // nous mettons se dépôt dans la liste des utilisables
   if (DS.State=dsInsert) and ( VH_GC.GCMultiDepots and GetParamSoc('SO_GCLIAISONAUTODEP_ETAB')) then
      begin
      if GetField('ET_DEPOTLIE')<>'' then
         SetField('ET_DEPOTLIE',GetField('ET_DEPOTLIE')+';'+GetField('ET_ETABLISSEMENT'));
      end;
   // 1 - Vérification que le dépôt correspondant à l'établissement existe
   //     Si ce n'est pas le cas, on le crée
   TOB_DEPOT:=TOB.Create('DEPOTS', NIL, -1);  // Création de la TOB des dépôts
   if not ExisteSQL('SELECT GDE_DEPOT FROM DEPOTS WHERE GDE_DEPOT="'+GetField('ET_ETABLISSEMENT')+'"') then
      begin
      TOB_DEPOT.PutValue('GDE_DEPOT',GetField('ET_ETABLISSEMENT')) ;
      TOB_DEPOT.PutValue('GDE_LIBELLE',GetField('ET_LIBELLE')) ;
      TOB_DEPOT.PutValue('GDE_ABREGE',GetField('ET_ABREGE')) ;
      TOB_DEPOT.PutValue('GDE_ADRESSE1',GetField('ET_ADRESSE1')) ;
      TOB_DEPOT.PutValue('GDE_ADRESSE2',GetField('ET_ADRESSE2')) ;
      TOB_DEPOT.PutValue('GDE_ADRESSE3',GetField('ET_ADRESSE3')) ;
      TOB_DEPOT.PutValue('GDE_CODEPOSTAL',GetField('ET_CODEPOSTAL')) ;
      TOB_DEPOT.PutValue('GDE_VILLE',GetField('ET_VILLE')) ;
      TOB_DEPOT.PutValue('GDE_PAYS',GetField('ET_PAYS')) ;
      TOB_DEPOT.PutValue('GDE_TELEPHONE',GetField('ET_TELEPHONE')) ;
      TOB_DEPOT.PutValue('GDE_FAX',GetField('ET_FAX')) ;
      TOB_DEPOT.PutValue('GDE_EMAIL',GetField('ET_EMAIL')) ;
      TOB_DEPOT.PutValue('GDE_UTILISATEUR',GetField('ET_UTILISATEUR')) ;
      if (DS.State=dsInsert) then TOB_DEPOT.PutValue('GDE_DATECREATION',GetField('ET_DATECREATION')) ;
      TOB_DEPOT.PutValue('GDE_DATEMODIF',GetField('ET_DATEMODIF')) ;
      TOB_DEPOT.PutValue('GDE_SURSITE',GetField('ET_SURSITE')) ;
      TOB_DEPOT.PutValue('GDE_SURSITEDISTANT',GetField('ET_SURSITEDISTANT')) ;
      TOB_DEPOT.InsertDB(nil,FALSE);//InsertOrUpdateDB (FALSE);
      end
   else
      begin     // modification du dépôt
      QQ:=OpenSQL('SELECT Count(*) FROM DEPOTS WHERE GDE_DEPOT="'+GetField('ET_ETABLISSEMENT')+'"',True) ;
      Nb:=QQ.Fields[0].AsInteger ;
      Ferme(QQ) ;
      if (Nb=1) and (DS.State=dsInsert) and ( VH_GC.GCMultiDepots and GetParamSoc('SO_GCLIAISONAUTODEP_ETAB')) then
            if HShowMessage('1;;Ce dépôt existe déjà, voulez-vous l''écraser?;W;YN;N;N;;;','','')=mrYes then
      if (Nb=1) and (TOB_DEPOT.SelectDB('"'+GetField('ET_ETABLISSEMENT')+'"',Nil)=TRUE) then
         BEGIN

            begin
            TOB_DEPOT.LoadDB(FALSE) ;
            TOB_DEPOT.PutValue('GDE_DEPOT',GetField('ET_ETABLISSEMENT')) ;
            TOB_DEPOT.PutValue('GDE_LIBELLE',GetField('ET_LIBELLE')) ;
            TOB_DEPOT.PutValue('GDE_ABREGE',GetField('ET_ABREGE')) ;
            TOB_DEPOT.PutValue('GDE_ADRESSE1',GetField('ET_ADRESSE1')) ;
            TOB_DEPOT.PutValue('GDE_ADRESSE2',GetField('ET_ADRESSE2')) ;
            TOB_DEPOT.PutValue('GDE_ADRESSE3',GetField('ET_ADRESSE3')) ;
            TOB_DEPOT.PutValue('GDE_CODEPOSTAL',GetField('ET_CODEPOSTAL')) ;
            TOB_DEPOT.PutValue('GDE_VILLE',GetField('ET_VILLE')) ;
            TOB_DEPOT.PutValue('GDE_PAYS',GetField('ET_PAYS')) ;
            TOB_DEPOT.PutValue('GDE_TELEPHONE',GetField('ET_TELEPHONE')) ;
            TOB_DEPOT.PutValue('GDE_FAX',GetField('ET_FAX')) ;
            TOB_DEPOT.PutValue('GDE_EMAIL',GetField('ET_EMAIL')) ;
            TOB_DEPOT.PutValue('GDE_UTILISATEUR',GetField('ET_UTILISATEUR')) ;
            TOB_DEPOT.PutValue('GDE_DATEMODIF',GetField('ET_DATEMODIF')) ;
            TOB_DEPOT.PutValue('GDE_SURSITE',GetField('ET_SURSITE')) ;
            TOB_DEPOT.PutValue('GDE_SURSITEDISTANT',GetField('ET_SURSITEDISTANT')) ;
            TOB_DEPOT.UpdateDB(FALSE)//InsertOrUpdateDB (FALSE);
            end;
         END;
      end;
   TOB_DEPOT.free;
   end;
//{$ENDIF}
{Etablis_sursite:=(GetField('ET_SURSITE')='X') ;
end;  }

// **************************Fonction pour gestion des dépôts liés ***************************************

procedure TOM_Etabliss.RefreshGrid(posListe, posUtilisable: integer);
begin
  TobUtilisable.PutGridDetail(GUTILISABLE, False, False, 'GDE_DEPOT;GDE_LIBELLE', True);
  TobListe.PutGridDetail(GLISTE, False, False, 'GDE_DEPOT;GDE_LIBELLE', True);
  GUTILISABLE.Row := Min(posUtilisable, GUTILISABLE.RowCount - 1);
  GLISTE.Row := Min(posListe, GLISTE.RowCount - 1);
  RefreshBouton;
end;

// Boutons enable / disable

procedure TOM_Etabliss.RefreshBouton;
begin
  if (Ecran <> nil) and (Ecran is TFFiche) and
    (TFFiche(Ecran).TypeAction = taConsult) then
  begin
    DisableBoutons;
  end else
  begin
    SetControlEnabled('BFLECHEDROITE', TobListe.Detail.Count > 0);
    SetControlEnabled('BFLECHEGAUCHE', TobUtilisable.Detail.Count > 0);
    SetControlEnabled('BFLECHEHAUT', GUTILISABLE.Row > 0);
    SetControlEnabled('BFLECHEBAS', GUTILISABLE.Row < GUTILISABLE.RowCount - 1);
    SetControlEnabled('BFLECHETOUS', TobListe.Detail.Count > 0);
    SetControlEnabled('BFLECHEAUCUN', TobUtilisable.Detail.Count > 0);
  end;
end;

procedure TOM_Etabliss.DisableBoutons;
begin
  SetControlEnabled('BFLECHEDROITE', False);
  SetControlEnabled('BFLECHEGAUCHE', False);
  SetControlEnabled('BFLECHEHAUT', False);
  SetControlEnabled('BFLECHEBAS', False);
  SetControlEnabled('BFLECHETOUS', False);
  SetControlEnabled('BFLECHEAUCUN', False);
  SetControlEnabled('BINSERTDEP', False);
end;

procedure TOM_Etabliss.ClickFlecheDroite;
var indiceFille: integer;
begin
  // Y a t il quelque chose de sélectionné ?
  if GLISTE.Row < 0 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if TobUtilisable.Detail.Count > 0 then indiceFille := GUTILISABLE.Row + 1 else indiceFille := 0;
  TobListe.detail[GLISTE.Row].ChangeParent(TobUtilisable, indiceFille);
  RefreshGrid(GLISTE.Row, GUTILISABLE.Row + 1);
end;

procedure TOM_Etabliss.ClickFlecheGauche;
var indiceFille: integer;
begin
  // Y a t il quelque chose de sélectionné ?
  if GUTILISABLE.Row < 0 then exit;
  // Changement du parent de l'élément des établissements affichés
  if TobListe.Detail.Count > 0 then indiceFille := GLISTE.Row + 1 else indiceFille := 0;
  TobUtilisable.detail[GUTILISABLE.Row].ChangeParent(TobListe, indiceFille);
  RefreshGrid(GLISTE.Row + 1, GUTILISABLE.Row);
end;

procedure TOM_Etabliss.ClickFlecheTous;
var indiceFille, iGrd, posListe: integer;
begin
  if GLISTE.RowCount < 1 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if GUTILISABLE.RowCount > 1 then indiceFille := GUTILISABLE.Row + 1 else indiceFille := 0;
  posListe := TobListe.detail.count - 1;
  for iGrd := 0 to posListe do TobListe.detail[0].ChangeParent(TobUtilisable, indiceFille + iGrd);
  RefreshGrid(0, indiceFille + posListe);
end;

procedure TOM_Etabliss.ClickFlecheAucun;
var indiceFille, iGrd, posUtilisable: integer;
begin
  if GUTILISABLE.RowCount < 1 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if GLISTE.RowCount > 1 then indiceFille := GLISTE.Row + 1 else indiceFille := 0;
  posUtilisable := TobUtilisable.detail.count - 1;
  for iGrd := 0 to posUtilisable do TobUtilisable.detail[0].ChangeParent(TobListe, indiceFille + iGrd);
  RefreshGrid(indiceFille + posUtilisable, 0);
end;

procedure TOM_Etabliss.ClickFlecheHaut;
begin
  if GUTILISABLE.Row < 1 then exit;
  // Changement de l'indice dans la Tob parent
  TobUtilisable.detail[GUTILISABLE.Row].ChangeParent(TobUtilisable, GUTILISABLE.Row - 1);
  RefreshGrid(GLISTE.Row, GUTILISABLE.Row - 1);
end;

procedure TOM_Etabliss.ClickFlecheBas;
begin
  if GUTILISABLE.Row > GUTILISABLE.RowCount - 2 then exit;
  // Changement de l'indice dans la Tob parent
  TobUtilisable.detail[GUTILISABLE.Row].ChangeParent(TobUtilisable, GUTILISABLE.Row + 1);
  RefreshGrid(GLISTE.Row, GUTILISABLE.Row + 1);
end;

procedure TOM_Etabliss.ClickInsertDepot(Depot: string);
var
  QQ: TQuery;
begin
  // Ajout du dépot créer dans la liste des utilisables
  QQ := OpenSQL('select GDE_DEPOT,GDE_LIBELLE,GDE_ABREGE from DEPOTS where GDE_DEPOT="' + Depot + '"', True);
  TobUtilisable.LoadDetailDB('DEPOTS', '', '', QQ, True);
  Ferme(QQ);
  RefreshGrid(GLISTE.Row, GUTILISABLE.Row + 1);
end;

procedure TOM_Etabliss.ValideListeDepot;
var QQ: TQuery;
  i_ind, NbMaxDepot: integer;
  ListeDepot: string;
begin
  // Construction de la liste des dépots lié
  // Dans le cas des multi-dépôts
  // Si tous les dépots sont selectionnés alors je ne met rien dans le champ depots liés
  if GetParamSoc('SO_GCMULTIDEPOTS') then
  begin
    QQ := OpenSQL('SELECT Count(*) FROM DEPOTS', True);
    NbMaxDepot := QQ.Fields[0].AsInteger;
    Ferme(QQ);
    ListeDepot := '';
    if TobUtilisable.Detail.count <> NbMaxDepot then
    begin
      for i_ind := 0 to TobUtilisable.Detail.count - 1 do
      begin
        if ListeDepot <> '' then ListeDepot := ListeDepot + ';';
        ListeDepot := ListeDepot + TobUtilisable.detail[i_ind].GetValue('GDE_DEPOT');
      end;
    end;
    DS.edit;
    SetField('ET_DEPOTLIE', ListeDepot);
  end;
end;

// *********************************Fonction AGL ***********************************************************

procedure AGLControleAffPhoto_ET(parms: array of variant; nb: integer);
{$IFDEF GCGC}
{$IFNDEF PGIMAJVER}
var F: TForm;
  OM: TOM;
  {$ENDIF PGIMAJVER}
  {$ENDIF}
begin
  {$IFDEF GCGC}
  {$IFNDEF PGIMAJVER}
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Etabliss) then TOM_Etabliss(OM).AffichePhoto else exit;
  {$ENDIF PGIMAJVER}
  {$ENDIF}
end;
// Procedure permettant d'affecter la liste des depôts liés dans ET_DEPOTLIE

procedure AGLValideListeDepot(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Etabliss) then TOM_Etabliss(OM).ValideListeDepot else exit;
end;

// Permet de gerer les boutons de navigation des dépôts d'une liste à l'autre

procedure AGLOnClickBouton1(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Etabliss) then
  begin
    if TFFiche(F).TypeAction = taConsult then Exit;
    if Parms[1] = BTN_DROIT then TOM_Etabliss(OM).ClickFlecheDroite
    else if Parms[1] = BTN_GAUCHE then TOM_Etabliss(OM).ClickFlecheGauche
    else if Parms[1] = BTN_HAUT then TOM_Etabliss(OM).ClickFlecheHaut
    else if Parms[1] = BTN_BAS then TOM_Etabliss(OM).ClickFlecheBas
    else if Parms[1] = BTN_TOUS then TOM_Etabliss(OM).ClickFlecheTous
    else if Parms[1] = BTN_AUCUN then TOM_Etabliss(OM).ClickFlecheAucun
    else if Parms[1] = BTN_INSERT then TOM_Etabliss(OM).ClickInsertDepot(Parms[2])
    else if (Parms[1] = GRD_LISTE) or (Parms[1] = GRD_UTILISABLE) then TOM_Etabliss(OM).RefreshBouton;
  end;
end;

procedure TOM_Etabliss.TPOnEnter (Sender:TObject); //LM20070201
var s : string ;
begin
  if sender=nil then exit ;
  s := TControl(Sender).name ;
  NTP_.AddChampSupValeur(s, getControlText(s));
end ;

procedure TOM_Etabliss.TPOnExit (Sender:TObject); //LM20070201
var s:string;
begin
  if sender=nil then exit ;
  s := TControl(Sender).name ;
  if NTP_.g(s)<>getControlText(s) then
  begin
    if (ds.state in [dsEdit, dsInsert]) then else ForceUpdate;
  end ;

end ;

//DEB FC20070904
{$IFDEF PAIEGRH}
procedure Tom_Etabliss.DupliquerEtab(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
{$ELSE}
  Code: THEdit;
{$ENDIF}
  AncValCode, AncValLib: string;
  Champ: array[1..1] of Hstring;
  Valeur: array[1..1] of variant;
  Ok: Boolean;
  st: string;
  TOB_EtabCompl,T_Dupli,T : Tob;
  i:integer;
begin
  Dupli := True;
  TFFiche(Ecran).BValider.Click;
  AncValCode := GetField('ET_ETABLISSEMENT');
  AncValLib := GetField('ET_LIBELLE');
  AglLanceFiche('PAY', 'DUPLI_ETAB', '', '', AncValCode + ';' + AncValLib);

  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'ET_ETABLISSEMENT';
    Valeur[1] := PGCodeDupliquer;
    Ok := RechEnrAssocier('ETABLISS', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
      TOB_EtabCompl := TOB.Create('t_etabcompl', nil, -1);
      st := 'SELECT * FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="' + AncValCode + '"';
      TOB_EtabCompl.LoadDetailDBFROMSQL('ETABCOMPL', St);
      T_Dupli := TOB.Create('etabcompl dupliqué', nil, -1);
      DupliquerPaie(TFFicheListe(Ecran).TableName, Ecran);
      SetField('ET_ETABLISSEMENT', PGCodeDupliquer);
      for i := 0 to TOB_EtabCompl.Detail.Count - 1 do
      begin
        if ((TOB_EtabCompl.Detail[i].GetValue('ETB_ETABLISSEMENT')) = AncValCode) then
        begin
          T := TOB_EtabCompl.Detail[i];
          if T <> nil then
          begin
            T.PutValue('ETB_ETABLISSEMENT', PGCodeDupliquer);
            T.PutValue('ETB_LIBELLE', PGLibDupliquer);
          end;
        end;
      end;
      T_Dupli.Dupliquer(TOB_EtabCompl, TRUE, TRUE, FALSE);
      T_Dupli.InsertDB(nil, False);
      TOB_EtabCompl.free;
      T_Dupli.free;
      SetField('ET_ETABLISSEMENT', PGCodeDupliquer);
      SetField('ET_LIBELLE', PGLibDupliquer);

      SetControlEnabled('BInsert', True);
      SetControlEnabled('ET_ETABLISSEMENT', False);
      SetControlEnabled('BDUPLIQUER', True);
      TFFicheListe(Ecran).Bouge(nbPost); //Force enregistrement
    end
    else
      HShowMessage('5;Etablissement :;La duplication est impossible, l''établissement existe déjà.;W;O;O;O;;;', '', '');
  end;
  Dupli := False;
end;
{$ENDIF PAIEGRH}
//FIN FC20070904
initialization
  registerclasses([TOM_Etabliss]);
  RegisterAglProc('AffichePhotoEtablissement', TRUE, 0, AGLControleAffPhoto_ET);
  RegisterAglProc('OnClickBouton1', True, 1, AGLOnClickBouton1);
  RegisterAglProc('ValideListeDepot', True, 1, AGLValideListeDepot);
end.

