{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : HISTOAGENDA ()
Mots clefs ... : TOF;HISTOAGENDA
*****************************************************************}
Unit utofhistoagenda ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     maineagl,
     db,
     dbtables,
{$ELSE}
     mul,
     fe_main,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HDB,
     HMsgBox,
     UTOF,
     DPTofAnnuSel
     ;

Type
  TOF_HISTOAGENDA = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    procedure OnDblClickListe     (Sender:TObject);
    procedure OnClickRechAnnuaire (Sender:TObject);
    procedure OnClickRechMission  (Sender:TObject);

    procedure OnNewActInterne     (Sender:TObject); // $$$ JP 24/04/06
    procedure OnNewActExterne     (Sender:TObject); // $$$ JP 24/04/06

    // $$$ JP 29/03/07: multimaj
    procedure OnFait        (Sender:TObject);
    procedure OnNoAlerte    (Sender:TObject);

  private
         m_TOBDroits   :TOB;

         function  GetSelectedGuids (strMessage:string):string;
         procedure NewActivite (bExterne:boolean);

  end ;

procedure DPLance_HistoAgenda;
procedure DPLance_MultiMajAgenda; // $$$ JP 29/03/07

Implementation

uses
  dpoutilsagenda,
  UDossierSelect,
  utofaffaire_mul,
  menus,
  AGLInit, paramsoc,
  hqry;

procedure DPLance_HistoAgenda;
begin
     if VH_Doss <> nil then
        AGLLanceFiche ('YY', 'YYHISTOAGENDA', '', '', '');
end;

procedure DPLance_MultiMajAgenda;
begin
     AGLLanceFiche ('YY', 'YYHISTOAGENDA', '', '', 'MULTIMAJ');
end;



procedure TOF_HISTOAGENDA.OnNewActInterne (Sender:TObject); // $$$ JP 24/04/06
begin
     NewActivite (FALSE);
end;

procedure TOF_HISTOAGENDA.OnNewActExterne (Sender:TObject); // $$$ JP 24/04/06
begin
     NewActivite (TRUE);
end;

procedure TOF_HISTOAGENDA.NewActivite (bExterne:boolean);
var
   strArg   :string;
   strUsers :string;
begin
     // Ajout des seuls codes utilisateur possible en saisie
     strUsers := AgendaGetUsers (V_PGI.User, m_TOBDroits, bExterne, FALSE);
     if strUsers = '' then
     begin
          PgiInfo ('Vous n''avez pas les droits suffisants pour créer une activité de ce type');
          exit;
     end;

     strArg := 'ACTION=CREATION;USERS=' + strUsers;
     if bExterne = TRUE then
         strArg := strArg + ';JEV_EXTERNE=X'
     else
         strArg := strArg + ';JEV_EXTERNE=-';
     strArg := strArg + ';JEV_ABSENCE=-';
     strArg := strArg + ';JEV_USER1=' + V_PGI.User;

     // Dossier en cours
     strArg := strArg + ';JEV_NODOSSIER=' + VH_DOSS.NoDossier;

     // Lancement en création de la fiche agenda
     AgendaLanceFiche ('', strArg);
     AGLRefreshDB ([LongInt (Ecran), 'FListe'], 2);
end;

procedure TOF_HISTOAGENDA.OnDblClickListe (Sender:TObject);
var
   strArg        :string;
   TOBAct        :TOB;
   strGuidEvt    :string;
   iAutorisation :integer;
   strUsers      :string;
begin
     // $$$ JP 24/04/06 - si les droits sont suffisants, on autorise la mise à jour
     with TFMul (Ecran).Q do
     begin
          strGuidEvt := FindField ('JEV_GUIDEVT').AsString;
          if strGuidEvt <> '' then
          begin
               TOBAct := TOB.Create ('activite agenda', nil, -1);
               try
                  TOBAct.LoadDetailFromSQL ('SELECT JEV_USER1 AS I_CODE,JEV_EXTERNE AS I_EXTERNE,JEV_ABSENCE AS I_ABSENCE,JEV_PERS AS I_PERS,JEV_FAIT AS I_FAIT,YX_LIBRE AS I_DROIT FROM JUEVENEMENT LEFT JOIN CHOIXEXT ON (YX_TYPE="DAU" AND YX_LIBELLE="' + V_PGI.User + '" AND YX_ABREGE=JEV_USER1) WHERE JEV_GUIDEVT="' + strGuidEvt + '"');
                  if TOBAct.Detail.Count > 0 then
                  begin
                       iAutorisation := AgendaMajType (TOBAct.Detail [0]);
                       if iAutorisation <> aaAll then
                           strArg := 'ACTION=CONSULTATION'
                       else
                           strArg := 'ACTION=MODIFICATION';
                       if iAutorisation <> aaNothing then
                          strArg := strArg + ';MODIFAUTORISEE';
                  end
                  else
                      strArg := 'ACTION=CONSULTATION';
               finally
                      TOBAct.Free;
               end;

               // $$$ JP 24/04/06 - utilisateurs utilisables par l'utilisateur en cours
               strUsers := AgendaGetUsers (V_PGI.User, m_TOBDroits, GetField ('JEV_EXTERNE')='X', GetField ('JEV_ABSENCE')='X', GetField ('JEV_USER1'));
               if strUsers <> '' then
                  strArg := strArg + ';USERS=' + strUsers;

               //strArg := 'ACTION=CONSULTATION';
               strArg := strArg + ';JEV_EXTERNE=' + FindField ('JEV_EXTERNE').AsString;
               strArg := strArg + ';JEV_ABSENCE=' + FindField ('JEV_ABSENCE').AsString;
               AgendaLanceFiche (FindField ('JEV_GUIDEVT').AsString, strArg);
               AGLRefreshDB ([LongInt (Ecran), 'FListe'], 2);
          end;
     end;
end;

procedure TOF_HISTOAGENDA.OnClickRechAnnuaire(Sender: TObject);
var
   sGuidPer   :string;
   TOBLibelle   :TOB;
begin
     sGuidPer := LancerAnnuSel ('', '', '');
     if sGuidPer <> '' then
     begin
          SetControlText ('JEV_GUIDPER', sGuidPer);

          TOBLibelle := TOB.Create ('le libelle', nil, -1);
          try
             TOBLibelle.LoadDetailFromSQL ('SELECT ANN_NOM1 FROM ANNUAIRE WHERE ANN_GUIDPER="' + SGuidPer+'"');
             if TOBLibelle.Detail.Count > 0 then
                 SetControlText ('EGUIDPER', TOBLibelle.Detail [0].GetValue ('ANN_NOM1'))
             else
                 SetControlText ('EGUIDPER', '');
          finally
                 TOBLibelle.Free;
          end;
     end;
end;

procedure TOF_HISTOAGENDA.OnClickRechMission (Sender:TObject);
var
   strCodeMission   :string;
   TOBTiers         :TOB;
begin
     // Lien vers tiers associé au code dossier sélectionné (pas code personne, car peut être non lié)
     TOBTiers := TOB.Create ('le tiers', nil, -1);
     try
        TOBTiers.LoadDetailFromSQL ('SELECT T_TIERS FROM TIERS JOIN ANNUAIRE ON T_TIERS=ANN_TIERS JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="' + VH_Doss.NoDossier + '"');
        if TOBTiers.Detail.Count > 0 then
        begin
             // $$$JP 27/11/03: pour utiliser la tof affaire
             strCodeMission := AFLanceFiche_AffaireRech ('', 'AFF_TIERS=' + TOBTiers.Detail [0].GetValue ('T_TIERS'));
             if strCodeMission <> '' then
             begin
                  strCodeMission := ReadTokenSt (strCodeMission);
                  SetControlText ('JEV_AFFAIRE', strCodeMission);
                  strCodeMission := ReadTokenSt (strCodeMission);
                  SetControlText ('EMISSION', strCodeMission);
             end;
        end
        else
            PgiInfo ('Aucun tiers associé au dossier ' + VH_Doss.NoDossier + ' dans la base cabinet');
     finally
            TOBTiers.Free;
     end;
end;

procedure TOF_HISTOAGENDA.OnArgument (S:string);
var
  strEnabledUsers  :string;
  strWhere         :string;
begin
     inherited ;

     // $$$ JP 24/04/06 - Droits agenda
     m_TOBDroits := TOB.Create ('droits agenda', nil, -1);
     AgendaLoadDroits (m_TOBDroits);

     // vérification que des types d'activité existe
     if ExisteSQL ('SELECT JTE_CODEEVT FROM JUTYPEEVT WHERE JTE_FAMEVT="ACT"') = FALSE then
        PgiInfo ('Veuillez contacter l''administrateur Cegid Expert pour paramétrer les types d''activité (module Administration/Agenda)');

     SetControlProperty ('JEV_CODEEVT', 'Plus', 'AND JTE_FAMEVT="ACT"');

     // $$$ JP 29/03/07: pour maj multiple, pas de filtre dossier sinon pas d'activité récurrente
     if S<>'MULTIMAJ' then
     begin
          strWhere := 'JEV_NODOSSIER="' + VH_Doss.NoDossier + '"';
          strEnabledUsers := AgendaGetUsersMinRights (V_PGI.User, m_TOBDroits);

          // $$$ JP 23/05/07
          SetControlVisible ('BMULTIMAJ', FALSE);
     end
     else
     begin
          // $$$ JP 29/03/07: dans ce cas, il faut aussi enlever les activités sur lesquelles l'utilisateur n'a pas tout les droits
          strWhere := 'JEV_OCCURENCEEVT<>"REC"';
          strEnabledUsers := AgendaGetUsersMaxRights (V_PGI.User, m_TOBDroits);

          // $$$ JP 11/04/07: ne pas voir le bouton "nouveau"
          SetControlVisible ('BINSERT', FALSE);
     end;

     if strEnabledUsers <> '' then
         strWhere := strWhere + ' AND JEV_USER1 IN (' + strEnabledUsers + ')'
     else
         strWhere := '0=1'; // $$$ JP 09/08/07 FQ 11616: pas d'agenda visible, donc il ne faut pas laisser le mul afficher quelque chose
     SetControlText ('XX_WHERE', strWhere);

     THEdit (GetControl('EGUIDPER')).OnElipsisClick := OnClickRechAnnuaire;
     THEdit (GetControl('EMISSION')).OnElipsisClick := OnClickRechMission;

{$IFDEF EAGLCLIENT}
     THGrid (GetControl('FLISTE')).OnDblClick        := OnDblClickListe;
{$ELSE}
     THDBGrid (GetControl('FLISTE')).OnDblClick      := OnDblClickListe;
{$ENDIF}

     // $$$ JP 24/04/06 - FQ 10094/10095
     TMenuItem (GetControl ('NEW_ACTINTERNE')).OnClick := OnNewActInterne;
     TMenuItem (GetControl ('NEW_ACTEXTERNE')).OnClick := OnNewActExterne;

     // $$$ JP 29/03/07: multimaj
     if S='MULTIMAJ' then
     begin
          // multimaj => selectall visible
          SetControlVisible ('bselectall', TRUE);

          // multimaj => multisélection
{$IFDEF EAGLCLIENT}
          THDBGrid(GetControl('FListe')).MultiSelect := True;
{$ELSE}
          THDBGrid(GetControl('FListe')).MultiSelection := True;
{$ENDIF}

          TMenuItem (GetControl ('MMAJ_FAIT')).OnClick     := OnFait;
          TMenuItem (GetControl ('MMAJ_NOALERTE')).OnClick := OnNoAlerte;
     end;
end;

procedure TOF_HISTOAGENDA.OnClose ;
begin
     inherited ;

     // $$$ JP 24/04/06
     FreeAndNil (m_TOBDroits);
end ;

function TOF_HISTOAGENDA.GetSelectedGuids (strMessage:string):string;
var
    Lst       :THDBGrid;
    iNbSel, i :integer;
    Q         :THQuery;
begin
     Result := '';

     Lst := TFMul(Ecran).FListe;
     Q   := TFMul(Ecran).Q;
     if (Lst.AllSelected = TRUE) or (Lst.NbSelected > 0) then
     begin
          if PGIAsk (strMessage, TitreHalley) <> mrYes then
            exit;

          // Traitement
          if Lst.AllSelected then
          begin
{$IFDEF EAGLCLIENT}
                if not TFMul(Ecran).FetchLesTous then
                   PGIInfo ('Impossible de récupérer tous les enregistrements')
                else
{$ENDIF}
                begin
                     Q.First;
                     while Not Q.EOF do
                     begin
                          if Result <> '' then
                             Result := Result + '","';
                          Result := Result + Q.FindField ('JEV_GUIDEVT').AsString;

                          Q.Next;
                     end;
                end;
          end
          else
          begin
               iNbSel := Lst.NbSelected;
               for i := 0 to iNbSel-1 do
               begin
                    Lst.GotoLeBookmark (i);
{$IFDEF EAGLCLIENT}
                    Q.TQ.Seek(Lst.Row - 1) ;
{$ENDIF}
                    if Result <> '' then
                       Result := Result + '","';
                    Result := Result + Q.FindField ('JEV_GUIDEVT').AsString;
               end;
          end;
     end;
end;

procedure TOF_HISTOAGENDA.OnFait(Sender: TObject);
var
   strGuids :string;
begin
     // Màj fait sur les guids sélectionnés
     strGuids := GetSelectedGuids ('Confirmez-vous le statut REALISE pour toutes les activités sélectionnées?');
     if strGuids <> '' then
        if GetParamSocSecur ('SO_AGEGENEREGI', True) = TRUE then
            AgendaSendToGi ('"' + strGuids + '"', TRUE)
        else
            ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_FAIT="X",JEV_DATEMODIF="' + UsDateTime (Now) + '",JEV_UTILISATEUR="' + V_PGI.User + '" WHERE JEV_FAIT<>"X" AND JEV_GUIDEVT IN ("' + strGuids + '")');

     // Rafraichissement
     AGLRefreshDB ([LongInt(Ecran), 'FListe'], 2);
end;

procedure TOF_HISTOAGENDA.OnNoAlerte(Sender: TObject);
var
   strGuids :string;
begin
     // Màj fait sur les guids sélectionnés
     strGuids := GetSelectedGuids ('Confirmez-vous la suppression de l''alerte pour toutes les activités sélectionnées?');
     if strGuids <> '' then
        ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_ALERTE="-",JEV_ALERTEDATE="01/01/1900",JEV_DATEMODIF="' + UsDateTime (Now) + '",JEV_UTILISATEUR="' + V_PGI.User + '" WHERE JEV_ALERTE="X" AND JEV_GUIDEVT IN ("' + strGuids + '")');

     // Rafraichissement
     AGLRefreshDB ([LongInt(Ecran), 'FListe'], 2);
end;


Initialization
  registerclasses ( [ TOF_HISTOAGENDA ] ) ;

end.

