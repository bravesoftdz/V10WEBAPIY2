{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Saisie des congés payés en mode groupé. cad pour un
Suite ........ : ensemble de salarié
Mots clefs ... : PAIE;CP
*****************************************************************
PT1   : 10/01/2002 SB V571 Affectation de ValidResp et ExportOk
PT2-1 : 07/02/2003 SB V591 FQ 10485 Affectation du Code etape
PT2-2 : 10/02/2003 SB V595 FQ 10485 Optimisation du traitement
PT3-1 : 09/10/2002 SB V_42 FQ 10883 Intégration des raccourçis clavier
                           calendrier
PT3-2 : 09/10/2002 SB V_42 Construction libellé idem saisie CP
PT4   : 11/12/2003 SB V_50 FQ 10984 exclusion des salariés non entrées ou
                           sorties
PT5   : 08/04/2004 SB V_50 FQ 11136 Ajout Gestion des congés payés niveau
                           salarié
PT6   : 28/04/2004 SB V_50 FQ 10984 Erreur comparaison
PT7   : 12/05/2004 SB V_50 FQ 11264 Identification de saisie groupée
PT8   : 14/05/2004 SB V_50 FQ 11167 Mise en place Multisélection, Balayage de la
                           Query du mul
PT9   : 14/10/2004 PH V_50 FQ 11710 Suppression de la multi sélection
PT10  : 24/01/2004 SB V_60 Refonte Saisie grille
                           FQ 11367 Calcul du nombre d'heure
                           FQ 11624 Saisie Groupée des absences
                           FQ 11854 Saisie Date de validité, début et fin de
                                    journée
                           FQ 11898 Permettre à nouveau la validation
PT11  : 01/07/2005 SB V_65 FQ 12386 Grammaire
PT12  : 27/09/2005 SB V_65 FQ 12493 Initialisation par défaut en ajout de ligne
PT13  : 06/12/2005 SB V_65 FQ 12692 Modification de la tablette des absences
PT14  : 07/04/2006 SB V_65 FQ 12941 Traitement nouveau mode de décompte motif
                           d'absence
PT15  : 11/04/2006 MF V_65 maj champs PCN_GESTIONIJSS, PCN_NBJCALEND,PCN_NBJIJSS
                           et PCN_NBJCARENCE : On tient compte de l'historique
                           des absences de type IJSS
PT16  : 13/04/2006 SB V_65 Traitement gestion maximum
PT17  : 13/04/2006 SB V_65 Traitement de l'annulation de l'absence
PT18  : 14/04/2006 SB V_65 FQ 11953 Paramsoc absence à cheval sur plusieurs mois
PT19  : 19/05/2006 SB V_65 FQ 13157 Refonte calcul no_ordre
PT20  : 01/06/2006 SB V_65 FQ 13032 Anomalie calcul heure
PT21  : 01/06/2006 SB V_65 Ajout du libellé de l'absence
PT22  : 09/06/2006 SB V_65 FQ 13202 Ajout option contrôle absence sur motif
                           d'absence
PT23  : 10/08/2007 FC V_72 FQ 13548
PT24  : 17/10/2007 VG V_80 Mise à jour du planning unifié
PT25  : 30/10/2007 GGU V_80 Gestion des absences en horaires
}
unit UTOFCongesGRP;

interface
uses  StdCtrls,
      Controls,
      Classes,
      sysutils,
      Windows,
{$IFNDEF EAGLCLIENT}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
      Hqry,
{$ENDIF}
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOF,
      UTOB,
      HTB97,
      ParamDat,
      Vierge,
      Ed_tools,
      UTOFCongesGrpMul,
      PgPlanningUnifie;

Type
     TOF_CongesGrp = Class (TOF)
       procedure OnArgument (stArgument : String ) ;     override ;
       procedure OnLoad ;                                override ;
       procedure OnUpdate ;                              override ;
       procedure OnClose;                                override ;
     private
       Grille                                            : THGrid;   { PT10 }
       Tob_Conges                                        : TOB;
       OkUpdate,SaisieLigne,DateError : Boolean;
       TheMul: TQUERY; // Query recuperee du mul  PT8
       StValue, StMode, StCaption : String;
       Plus : String;
       procedure MajCongespayesindividuel(TOB_Conges :TOB);
//       function  ConstruitStWhere( Prefixe : String) : string;
       Function  ExistCongesPris(Salarie: string; Tob_HistoPrisCp,T_PrisCp, Tob_MotifAbs : Tob; Var LibAbs : string) : Boolean; { PT21 PT22 } 
       procedure MiseEnFormeGrille(Row : integer; Vide : Boolean);                       { DEB PT10 }
       procedure AfficheValeurDefaut (Row : integer ; Vide : Boolean);
       procedure GrilleInsereLigne( sender : TObject );
       procedure GrilleDeleteLigne( sender : TObject );
       procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
       procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
       procedure GrilleCellKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
       procedure GrilleDblClick ( sender : TObject );
       procedure GrilleElipsisClick ( sender : TObject ); { FIN PT10 }
     END ;

implementation

uses PGoutils,PGCongesPayes,PGCommun,EntPaie, PgCalendrier, StrUtils;

procedure TOF_CongesGrp.OnArgument(stArgument: String);
var
Btn     : TToolBarButton97;
begin
Inherited ;
{ DEB PT10 }             
StMode := stArgument;
if StMode = 'CP' then StCaption := 'Saisie groupée des congés payés'
                 else StCaption := 'Saisie groupée des absences';
Ecran.Caption := StCaption;
UpdateCaption(Ecran);
SetControlVisible ('BSELECTALL', FALSE); // PT9
OkUpdate:=False; SaisieLigne:=False;
Grille:=THGrid(GetControl('GRILLE')) ;
if Grille<>Nil then
  begin
  MiseEnFormeGrille(1,False);
  Grille.OnElipsisClick := GrilleElipsisClick;
  Grille.OnKeyDown      := GrilleCellKeyDown;
  Grille.OnDblClick     := GrilleDblClick;
  Btn := TToolBarButton97(GetControl('BINS_LINE'));
  if btn<>nil then btn.OnClick := GrilleInsereLigne;
  Btn := TToolBarButton97(GetControl('BDEL_LINE'));
  if btn<>nil then btn.OnClick := GrilleDeleteLigne;
  Grille.OnCellEnter    := GrilleCellEnter;
  Grille.OnCellExit     := GrilleCellExit;
  End;
{ FIN PT10 }

  { PT8 Suppression de la partie servant à constituer une clause where bâti sur la sélection du mul }
  { DEB PT8 }
  if TFVierge(Ecran) <> nil then
  begin
    {$IFDEF EAGLCLIENT}
    TheMul := THQuery(TFVierge(Ecran).FMULQ).TQ;
    {$ELSE}
    TheMul := TFVierge(Ecran).FMULQ;
    {$ENDIF}
  end;
  { FIN PT8 }

end;


{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Procédure de mise à jour de la grille de saisie des absences 
Suite ........ : ou cp groupée
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
procedure TOF_CongesGrp.MajCongespayesindividuel(Tob_Conges:TOB);
var
QSt                                                              : tquery;
StrQuery,Salarie,Etab,SalarieAvant,LibAbs                        : string; { PT19 }
No_Ordre, i , j , Nodjm, Nodjp                                   : integer;
DateDebut, DateFin , DateEntree, DateSortie                      : tDateTime;
Tob_HistoPrisCp,Tob_CpPris,T_Pris,T,Tob_MotifAbs, T_MotifAbs     : Tob;   { PT14 }
TobCree,Error,OKValid                                            : Boolean;
ListeErreur                                                      : TListBox;
Nb_J, Nb_H                                                       : Double;
// d PT15
  WCarence: integer;
  Tob_AbsIJSS, AbsIJSS: TOB;
  WDateDeb, WDateFin: tdatetime;
  Q: tquery;
  st : string;
// f PT15
YYD, MMD, JJ, YYF, MMF                                           : WORD;
PGYPL : TypePGYPL;
begin
   { PT8 refonte procedure : balayage du mul pour génération des mvts CP pris }
   If TheMul = nil then exit;
   { Parcours du multicritère }
   TheMul.First;
   InitMoveProgressForm(nil, StCaption, 'Veuillez patienter SVP ...', TheMul.RecordCount, FALSE, TRUE); { PT10 }
   TobCree:=False;
   OkUpdate:=False;
   Tob_CpPris:=nil;
   No_ordre := 0; { PT10 }
   ListeErreur := TListBox (GetControl ('LBERREUR',True));
   ListeErreur.Items.Clear;

   Tob_MotifAbs := tob.create('tob_virtuelle', nil, -1);                         { PT14 }
   Tob_MotifAbs.loaddetaildb('MOTIFABSENCE', '', 'PMA_MOTIFABSENCE', nil, False);{ PT14 }

   { Bouclage sur le mul pour recup salariés concernés par la saisie groupée }
   while not TheMul.EOF do
   begin
      FreeAndNil(Tob_HistoPrisCp);
      i := 0; j :=0;   { PT10 }
      Salarie := TheMul.FindField('PSA_SALARIE').AsString;
      MoveCurProgressForm('Traitement du salarié : ' +Salarie+' '+RechDom('PGSALARIE',Salarie,False));
      { DEB PT2-2 Optimisation du traitement }
      { Chargement des absences déjà saisies }
      StrQuery:='SELECT PCN_SALARIE,PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_LIBELLE,PCN_TYPECONGE '+ { PT22 }
            'FROM ABSENCESALARIE '+
            'WHERE (PCN_TYPEMVT="ABS" OR PCN_TYPECONGE="PRI") AND PCN_SALARIE="'+Salarie+'" '+ //
            'AND PCN_ETATPOSTPAIE <> "NAN" '+ { PT17 }
            'ORDER BY PCN_ETABLISSEMENT';
      QSt:=OpenSql(StrQuery,True);
      if not QSt.eof then
        Begin
        Tob_HistoPrisCp:=TOB.Create('Les pri existants',Nil,-1) ;
        Tob_HistoPrisCp.loadDetailDb('Les pri existants','','',QSt,False);
        End;
      Ferme(QSt);
      // Balayage des lignes de congés à répartir
      T:=Tob_Conges.FindFirst([''],[''], false );
      while T <> nil do
         Begin
         Inc(j);                                                    { DEB PT10 }
         // on vérifie que la ligne est valide
         if ( (T.GetValue('DEBUTDJ')<>'') AND (T.GetValue('FINDJ')<>'') AND
              (T.GetValue('TYPECONGE')<>'') AND (T.GetValue('LIBELLE')<>'') AND
              isvaliddate(T.GetValue('DATEDEBUTABS')) and
              isvaliddate(T.GetValue('DATEFINABS'))   and
               ((T.GetValue('CALCUL')='OUI') OR
               ((T.GetValue('CALCUL')='NON') AND (T.GetValue('JOURS')<>'') AND (T.GetValue('HEURES')<>'')))  ) then
            begin
            DateDebut := strtodate(T.GetValue('DATEDEBUTABS'));
            DateFin   := strtodate(T.GetValue('DATEFINABS'));       { FIN PT10 }
            Error   := False;
            Etab    := TheMul.findfield('PSA_ETABLISSEMENT').AsString;
//            DTClot  := Q1.FindField('ETB_DATECLOTURECPN').AsDateTime;
            { DEB PT4 }
            DateEntree := TheMul.FindField('PSA_DATEENTREE').AsDateTime;
            DateSortie := TheMul.FindField('PSA_DATESORTIE').AsDateTime;
            if DateEntree>DateDebut then
               Begin
               ListeErreur.Items.Add('le salarié '+Salarie+' est entrée après la date de début d''absence : '+DateToStr(DateEntree)+'.');
               Error:=True;
               End;
            if (DateSortie>idate1900) and (DateSortie<DateFin) then { PT6 > au lieu de <> }
               Begin
               ListeErreur.Items.Add('le salarié '+Salarie+' est sorti avant la date de fin d''absence : '+DateToStr(DateSortie)+'.'); { PT6 }
               Error:=True;
               End;

            if Assigned(Tob_MotifAbs) then  { PT16 }
                 T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], [T.GetValue('TYPECONGE')], False);
                      { DEB PT10 }
            if (T.GetValue('TYPECONGE')='PRI') AND ((VH_Paie.PGCongesPayes = False) OR (RechDom('PGSALARIECONGES',Salarie,False) = ''))  then
               Begin
               ListeErreur.Items.Add('La gestion des congés payés n''est pas active pour le salarié : '+Salarie+'.');
               Error:=True;
               End
            else
               Begin
               AffecteNodemj(T.GetValue('DEBUTDJ'), Nodjm);
               AffecteNodemj(T.GetValue('FINDJ'), Nodjp);
               { DEB PT14 }
               //CalculDuree(DateDebut,DateFin,Salarie,Etab,T.GetValue('TYPECONGE'),Nb_J,Nb_H,Nodjm,Nodjp);
               CalculNbJourAbsence(DateDebut, DateFin, Salarie,
               Etab, T.GetValue('TYPECONGE'), T_MotifAbs, Nb_J, Nb_H, Nodjm, Nodjp);
               { FIN PT14 }
               if (Nb_J =0) and (T.GetValue('CALCUL')='OUI') then
                  Begin
                  ListeErreur.Items.Add('Le calendrier du salarié '+Salarie+' ne renvoie pas un nombre de jour d''absence positif!');
                  Error := True
                  End;
               End;   { FIN PT10 }

            { DEB PT16 Contôle Gestion maximun }
             if (T.GetValue('TYPECONGE')<>'PRI') then
               Begin
               If ControleGestionMaximum(Salarie,T.GetValue('TYPECONGE'),T_MotifAbs, DateDebut, Nb_J, Nb_H) then
                 Begin
                 if T_MotifAbs.GetValue('PMA_JOURHEURE') = 'JOU' then
                    begin
                    ListeErreur.Items.Add('Le nombre de jours maximun octroyés pour l''absence '+T.GetValue('LIBELLE')+' du salarié '+Salarie+' est dépassé.');
                    Error := True;
                    end
                 else
                 if (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU') or (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HOR') then //PT25
                   begin
                   ListeErreur.Items.Add('Le nombre d''heures maximun octroyés pour l''absence '+T.GetValue('LIBELLE')+' du salarié '+Salarie+' est dépassé.' );
                   Error := True;
                   end;
                 End;
               End;
             { FIN PT16 }

             { DEB PT18 Contôle absence à cheval }
             If not VH_Paie.PGAbsenceCheval then
                  Begin
                  DecodeDate(DateDebut,YYD,MMD,JJ);
                  DecodeDate(DateFin,YYF,MMF,JJ);
                  IF (MMD <> MMF) OR (YYD <> YYF) then
                    Begin
                    ListeErreur.Items.Add('L'' absence '+T.GetValue('LIBELLE')+' du salarié '+Salarie+' ne peut être à cheval sur plusieurs mois.');
                    Error := True;
                    End;
                  End;
             { FIN PT18 }
            if not Error then
{ BnotError }  Begin
               { FIN PT4 }
               { On intègre la condition dans la requête  : si etablissement gérant les CP }
               { On test que le congés pris n'est pas à cheval sur un autre mouvement }
               if not ExistCongesPris(Salarie,Tob_HistoPrisCp,T,Tob_MotifAbs,LibAbs) then { PT21 PT22 }
                  Begin
                  Inc(i);                                           { DEB PT10 }
                  { Optimisation du traitement utilisation d'une Tob }
                  { Creation de l'enregistrement dans la Table AbsenceSalarie }
                  if Tob_CpPris=nil then
                     Begin
                     Tob_CpPris:=TOB.Create('les absences groupée',nil,-1);  { PT19 }
                     if StMode = 'CP' then
                          No_ordre := IncrementeSeqNoOrdre('CPA',Salarie)
                     else
                        if T.GetValue('TYPECONGE')='PRI' then
                          No_ordre := IncrementeSeqNoOrdre('CPA',Salarie)
                        else
                          No_ordre := IncrementeSeqNoOrdre('',Salarie);
                     End
                  else
                    if SalarieAvant = Salarie then  { DEB PT19 }
                       Inc(No_ordre)
                  else
                  if SalarieAvant <> Salarie then  { FIN PT19 }
                     Begin
                     if StMode = 'CP' then
                          No_ordre := IncrementeSeqNoOrdre('CPA',Salarie) { PT19 }
                     else
                        if T.GetValue('TYPECONGE')='PRI' then
                          No_ordre := IncrementeSeqNoOrdre('CPA',Salarie) + i
                        else
                          No_ordre := IncrementeSeqNoOrdre('',Salarie) + i;
                     End ;                        { FIN PT10 }
                  SalarieAvant := Salarie; { PT19 }
                  TobCree:=True;
                  T_Pris:=TOB.Create('ABSENCESALARIE',Tob_CpPris,-1);
                  InitialiseTobAbsenceSalarie(T_Pris);   { DEB PT10 }

                  if T.GetValue('TYPECONGE')='PRI' then
                    Begin
                    T_Pris.PutValue('PCN_TYPEMVT'       , 'CPA');
                    T_Pris.PutValue('PCN_MVTPRIS'       , 'PRI')
                    End
                  else
                    Begin
                    T_Pris.PutValue('PCN_TYPEMVT'       , 'ABS');
                    T_Pris.PutValue('PCN_MVTPRIS'       , '');
                    End;                                 { FIN PT10 }
                  T_Pris.PutValue('PCN_SALARIE'       , Salarie);
                  T_Pris.PutValue('PCN_GUID'          , AglGetGUID()); //PT24
                  T_Pris.PutValue('PCN_ORDRE'         , No_Ordre);
                  T_Pris.PutValue('PCN_DATEDEBUTABS'  , DateDebut);
                  T_Pris.PutValue('PCN_DEBUTDJ'       , T.GetValue('DEBUTDJ'));  { PT10 }
                  T_Pris.PutValue('PCN_DATEFINABS'    , DateFin);
                  T_Pris.PutValue('PCN_FINDJ'         , T.GetValue('FINDJ'));    { PT10 }
                  T_Pris.PutValue('PCN_TYPECONGE'     , T.GetValue('TYPECONGE'));{ PT10 }
                  T_Pris.PutValue('PCN_SENSABS'       , '-');
                  if (T.GetValue('CALCUL')='OUI') then              { DEB PT10 }
                    Begin
                    T_Pris.PutValue('PCN_JOURS'       , Nb_j);
                    T_Pris.PutValue('PCN_HEURES'      , Nb_H);
                    End
                  else
                    Begin
                    T_Pris.PutValue('PCN_JOURS'       , T.getvalue('JOURS'));
                    T_Pris.PutValue('PCN_HEURES'      , T.getvalue('HEURES'));
                    End;                                            { FIN PT10 }
                  {T_Pris.PutValue('PCN_HEURES'        , floattostr(Nb_H)); PT20 }
                  T_Pris.PutValue('PCN_LIBELLE'       , T.getvalue('LIBELLE'));
                  T_Pris.PutValue('PCN_CODERGRPT'     , No_Ordre);
                  T_Pris.PutValue('PCN_DATEVALIDITE'  , StrToDate(T.GetValue('DATEVALIDITE'))); { PT10 }
                  T_Pris.PutValue('PCN_PERIODECP'     , 0);//CalculPeriode(DTClot, T.GetValue('DATEVALIDITE')));

                  T_Pris.PutValue('PCN_ETABLISSEMENT' , Etab);
                  T_Pris.PutValue('PCN_TRAVAILN1'     , TheMul.findfield('PSA_TRAVAILN1').asstring);
                  T_Pris.PutValue('PCN_TRAVAILN2'     , TheMul.findfield('PSA_TRAVAILN2').asstring);
                  T_Pris.PutValue('PCN_TRAVAILN3'     , TheMul.findfield('PSA_TRAVAILN3').asstring);
                  T_Pris.PutValue('PCN_TRAVAILN4'     , TheMul.findfield('PSA_TRAVAILN4').asstring);
                  T_Pris.PutValue('PCN_CODESTAT'      , TheMul.findfield('PSA_CODESTAT').asstring);
                  T_Pris.PutValue('PCN_CONFIDENTIEL'  , TheMul.findfield('PSA_CONFIDENTIEL').asstring);

                  T_Pris.PutValue('PCN_VALIDRESP'     , 'VAL');
                  T_Pris.PutValue('PCN_EXPORTOK'      , 'X');
                  T_Pris.PutValue('PCN_VALIDSALARIE'  , 'ADM');
                  T_Pris.PutValue('PCN_ETATPOSTPAIE'  , 'VAL');
                  T_Pris.PutValue('PCN_OKFRACTION'    , 'OUI');
                  T_Pris.PutValue('PCN_MVTORIGINE'    , 'GRP');  { PT7 }
// d PT15

                  if (T_MotifAbs.GetValue('PMA_GESTIONIJSS') = 'X') then
                  begin
                    T_Pris.PutValue('PCN_GESTIONIJSS','X');

                    Tob_AbsIJSS := Tob.Create('Absences IJSS', nil, -1);

                    st := 'SELECT PCN_SALARIE, PCN_DATEDEBUTABS, PCN_DATEFINABS, ' +
                          'PCN_LIBELLE, PCN_NBJCARENCE,PCN_GESTIONIJSS ' +
                          'FROM ABSENCESALARIE WHERE PCN_SALARIE = "' + Salarie +
                          '" AND PCN_GESTIONIJSS="X" AND ' +
                          'PCN_DATEDEBUTABS >="' + UsDateTime(PlusMois(DateDebut, -12)) + '" AND '+
                          'PCN_DATEDEBUTABS <="' + UsDateTime(DateDebut - 1) + '" ' +
                          'ORDER BY PCN_SALARIE, PCN_DATEFINABS';
                    Q := OpenSql(st, TRUE);

                    if not (Q.eof) then
                      Tob_AbsIJSS.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);

                    ferme(Q);

                    // détermination de la carence cumulée depuis 12 mois
                    WDateFin := IDate1900;
                    WCarence := 0;

                    AbsIJSS := Tob_AbsIJSS.FindFirst([''], [''], false);
                    while (AbsIJSS <> nil) do
                    begin
                      if (AbsIJSS.GetValue('PCN_DATEDEBUTABS') <> WDateFin + 1) then
                        WCarence := AbsIJSS.GetValue('PCN_NBJCARENCE')
                      else
                        WCarence := WCarence + AbsIJSS.GetValue('PCN_NBJCARENCE');

                      if (WCarence > T_MotifAbs.GetValue('PMA_CARENCEIJSS')) then
                        WCarence := T_MotifAbs.GetValue('PMA_CARENCEIJSS');

                      WDateFin := AbsIJSS.GetValue('PCN_DATEFINABS');
                      AbsIJSS := Tob_AbsIJSS.FindNext([''], [''], false);
                    end;

                    // calcul du nombre de jours calendaires d'absence
                    T_Pris.Putvalue('PCN_NBJCALEND', DateFin - DateDebut + 1);

                    // calcul du nbre de jours de carence
                    WDateDeb := T_Pris.Getvalue('PCN_DATEDEBUTABS');
                    if (WDateDeb = WDateFin + 1) then
                      T_Pris.Putvalue('PCN_NBJCARENCE', T_MotifAbs.GetValue('PMA_CARENCEIJSS') - WCarence)
                    else
                      T_Pris.Putvalue('PCN_NBJCARENCE', T_MotifAbs.GetValue('PMA_CARENCEIJSS'));

                    if (T_Pris.Getvalue('PCN_NBJCALEND') < T_Pris.Getvalue('PCN_NBJCARENCE')) then
                      T_Pris.Putvalue('PCN_NBJCARENCE', T_Pris.Getvalue('PCN_NBJCALEND'));

                    // calcul du nombre de jours d'IJSS
                    T_Pris.Putvalue('PCN_NBJIJSS',
                    T_Pris.Getvalue('PCN_NBJCALEND') - T_Pris.Getvalue('PCN_NBJCARENCE'));
                    FreeAndNil(Tob_AbsIJSS);

                  end;
// f PT15
                  End
               Else
                  ListeErreur.Items.Add('Absence existante "'+LibAbs+'" pour le salarié : '+Salarie); { PT10 } { PT21 }
{Endnoterror}  End;
            end
            else
            Begin                                                   { DEB PT10 }
            if (T.GetValue('DEBUTDJ')<>'') OR (T.GetValue('FINDJ')<>'') OR
                  (T.GetValue('TYPECONGE')<>'') OR (T.GetValue('LIBELLE')<>'') then
                 ListeErreur.Items.Add('La saisie de la ligne d''absence n°'+IntToStr(j)+' est incomplète!')
            else
              iF (T.GetValue('CALCUL')='NON') AND (Valeur(T.GetValue('JOURS'))<=0)   then
                 ListeErreur.Items.Add('Le nombre de jours de l''absence du '+T.GetValue('LIBELLE')+' doit être positif!')
              else
                iF (T.GetValue('CALCUL')='NON') AND (Valeur(T.GetValue('HEURES'))<=0)   then
                  ListeErreur.Items.Add('Le nombre d''heures de l''absence du '+T.GetValue('LIBELLE')+' doit être positif!')
                else
                  ListeErreur.Items.Add('Le format date de l''absence '+T.GetValue('LIBELLE')+' est incorrect!');
            End;                                                    { FIN PT10 }
         T:=Tob_Conges.FindNext([''] ,[''],false );
         end; //while T <> nil do
     TheMul.Next;
     End;

  if TobCree then
    Begin
    OKValid := True;
    if ListeErreur.Items.Count>0 then    { PT10 }
        if PgiAsk('Attention, vous avez des anomalies non bloquantes.#13#10'+ { PT11 }
                  'Confirmez-vous la mise à jour?',StCaption) = MrNo  then OKValid := False;
    if OkValid then
      try
      BeginTrans;
      MoveCurProgressForm('Mise à jour des données...');
      if Tob_CpPris<>nil then
        if Tob_CpPris.detail.count>0 then
//PT24
           begin
           Tob_CpPris.InsertDB(nil); { PT19 }
           for i:= 0 to Tob_CpPris.Detail.Count-1 do
               CreatePGYPL (Tob_CpPris.Detail[i], PGYPL);
           end;
//FIN PT24
      CommitTrans;
      PGIBox('La validation des absences groupées s''est bien effectuée.',StCaption); { PT10 }
      OkUpdate:=True;
      Except
      Rollback;
      PGIBox ('Une erreur est survenue lors de l''écriture des lignes d''absence!', StCaption); { PT10 }
      End
   else
     PGIBox ('Aucune ligne d''absence à valider!', StCaption); { PT6 12/05/2004 }
    End
  else
     PGIBox ('Aucune ligne d''absence à valider!', StCaption); { PT6 12/05/2004 }
FreeAndNil(Tob_CpPris);
FreeAndNil(Tob_HistoPrisCp);
FreeAndNil(T_MotifAbs);   { PT14 }
FreeAndNil(Tob_MotifAbs); { PT14 }
FiniMoveProgressForm;
{ FIN PT2-2 }
end;

procedure TOF_CongesGrp.OnLoad;
begin              // begin 1
//   GConges.Col:=0 ;GConges.Row:=1 ; PT3-1
   Tob_Conges:=TOB.Create('Congés groupés',Nil,-1) ;
   Tob_Conges.AddChampSup('TYPECONGE' ,false);        { DEB PT10 }
   Tob_Conges.AddChampSup('DATEDEBUTABS' ,false);
   Tob_Conges.AddChampSup('DEBUTDJ' ,false);
   Tob_Conges.AddChampSup('DATEFINABS' ,false);
   Tob_Conges.AddChampSup('FINDJ' ,false);
   Tob_Conges.AddChampSup('DATEVALIDITE' ,false);
   Tob_Conges.AddChampSup('CALCUL' ,false);
   Tob_Conges.AddChampSup('JOURS' ,false);
   Tob_Conges.AddChampSup('HEURES' ,false);
   Tob_Conges.AddChampSup('LIBELLE' ,false);          { FIN PT10 }
end;

{ DEB PT10 }
procedure TOF_CongesGrp.OnUpdate;
Var
i : integer;
begin
OkUpdate:=True;
if DateError then
  Begin
  PGiBox('Vous ne pouvez saisir deux absences à cheval sur une même période!',StCaption);
  Exit;
  End;
SaisieLigne := True;
For i := 1 to Grille.RowCount-1 do
  Begin
  if Not ( (Grille.cells[0,i]<>'')
    AND (IsValidDate(Grille.Cells[1,i])) AND (Grille.cells[2,i]<>'')
    AND (IsValidDate(Grille.Cells[3,i])) AND (Grille.cells[4,i]<>'')
    AND (IsValidDate(Grille.Cells[5,i]))
    AND (Grille.cells[9,i]<>'')
    AND ( (Grille.CellValues[6,i]='OUI')
       OR ((Grille.CellValues[6,i]='NON') AND (Grille.cells[7,i]<>'')
                                     AND (Grille.cells[8,i]<>''))) ) Then
      Begin
      SaisieLigne:=False;
      Break;
      End;
   End;
if (SaisieLigne=True) AND (Tob_Conges<>NIL) then //
  begin
  Tob_Conges.GetGridDetail(Grille,-1,'', 'TYPECONGE;DATEDEBUTABS;DEBUTDJ;DATEFINABS;FINDJ;DATEVALIDITE;CALCUL;JOURS;HEURES;LIBELLE');
  MajCongespayesindividuel(Tob_Conges);
  End
  else
  PGIBox('La saisie d''absence groupée n''est pas valide!',StCaption); { PT6 12/05/2004 }
end;
{ FIN PT10 }

{ DEB PT10 }
procedure TOF_CongesGrp.OnClose;
begin
if (OkUpdate=False) AND (SaisieLigne=True) then
  Begin
  If PgiAsk('Voulez-vous abandonner la saisie effectuée?',StCaption) = mrYes then
    Begin
    FreeAndNil(Tob_Conges);
    exit;
    End
  else
    begin
    LastErrorMsg :='' ;
    LastError:=1;
    end;
  End
else
    FreeAndNil(Tob_Conges);
end;
{ FIN PT10 }
(*function TOF_CongesGrp.ConstruitStWhere( Prefixe : String) : string;
begin
  { PT2-2 Contruction de la clause Where en fonction de la sélection multicritère }
  Result:='';
  if DeSalarie<>''     then Result:= 'AND '+Prefixe+'_SALARIE>="'+DeSALARIE+'" ';
  if ASalarie<>''      then Result:= Result+'AND '+Prefixe+'_SALARIE<="'+ASALARIE+'" ';
  if Etablissement<>'' then Result:= Result+'AND '+Prefixe+'_ETABLISSEMENT="'+Etablissement+'" ';
  if TravailN1<>''     then Result:= Result+'AND '+Prefixe+'_TRAVAILN1="'+TravailN1+'" ';
  if TravailN2<>''     then Result:= Result+'AND '+Prefixe+'_TRAVAILN2="'+TravailN2+'" ';
  if TravailN3<>''     then Result:= Result+'AND '+Prefixe+'_TRAVAILN3="'+TravailN3+'" ';
  if TravailN4<>''     then Result:= Result+'AND '+Prefixe+'_TRAVAILN4="'+TravailN4+'" ';
  if CodeStat<>''      then Result:= Result+'AND '+Prefixe+'_CODESTAT="'+CodeStat+'" ';
end;
            *)

  { PT2-2  recherche dans la tob des congés pris si mouvement existant }
Function TOF_CongesGrp.ExistCongesPris(Salarie: string; Tob_HistoPrisCp,T_PrisCp, Tob_MotifAbs : Tob; Var LibAbs : string) : Boolean; { PT21 }
Var T1,T2,T3 : Tob;
    AbsDebut,AbsFin : TDateTime;
    TCSaisie, TCBob : String;
begin
  Result:=False;
  if not assigned(Tob_HistoPrisCp) then exit;
  T1:=Tob_HistoPrisCp.FindFirst(['PCN_SALARIE'],[Salarie],False);
  While T1<>nil Do
    Begin
    AbsDebut := StrToDate(T_PrisCp.GetValue('DATEDEBUTABS'));
    AbsFin   := StrToDate(T_PrisCp.GetValue('DATEFINABS'));
    IF (AbsDebut>=T1.GetValue('PCN_DATEDEBUTABS')) AND (AbsDebut<=T1.GetValue('PCN_DATEFINABS')) then
         Result:=True;
    IF (AbsFin>=T1.GetValue('PCN_DATEDEBUTABS')) AND (AbsFin<=T1.GetValue('PCN_DATEFINABS')) then
       Result:=True;
    if Result then {DEB PT22 }
         begin
         If assigned(Tob_MotifAbs) and assigned(T_PrisCp) then
           Begin
           TCSaisie := T_PrisCp.GetValue('TYPECONGE');
           TCBob := T1.GetValue('PCN_TYPECONGE');
           if (TCSaisie <> TCBob) then
              Begin
              T2:=Tob_MotifAbs.findFirst(['PMA_MOTIFABSENCE'],[TCSaisie],False);
              IF Assigned(T2) then
                 if (T2.getValue('PMA_CONTROLMOTIF')='-')  then
                    Begin
                    T3 := Tob_MotifAbs.findFirst(['PMA_MOTIFABSENCE'],[TCBob],False);
                    if Assigned(T3) and (T3.getValue('PMA_CONTROLMOTIF') = 'X') then
                       result := False;
                    End
                    else
                 Result := False;
              End;
            End;  {FIN PT22 }  
         if Result then
           Begin
           LibAbs:=T1.GetValue('PCN_LIBELLE');
           Break;
           end;
         end; { PT21 }
    T1:=Tob_HistoPrisCp.FindNext(['PCN_SALARIE'],[Salarie],False);
    End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : Saisie groupée des absences, mise en forme de la grille
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.MiseEnFormeGrille(Row : integer; Vide : Boolean);
var
  //DEB PT23
  ValEtab : String;
  ValFinal : String;
  //FIN PT23
begin
  //DEB PT23
  Plus := '';
  ValFinal := '';
  if (CritEtab <> '') and (CritEtab <> '<<Tous>>') then
  begin
    // Remplacer les ; par des ,
    ValEtab := READTOKENST(CritEtab);
    while ValEtab <> '' do
    begin
      if ValFinal = '' then
        ValFinal := '"' + ValEtab + '"'
      else
        ValFinal := ValFinal + ',"' + ValEtab + '"';
      ValEtab := READTOKENST(CritEtab);
    end;
    if ExisteSQL('SELECT ETB_CONGESPAYES FROM ETABCOMPL WHERE ETB_CONGESPAYES="-" AND ETB_ETABLISSEMENT IN (' + ValFinal + ')') then
      Plus := ' AND PMA_MOTIFABSENCE <> "PRI"';
  end
  else
  begin
    if ExisteSQL('SELECT ETB_CONGESPAYES FROM ETABCOMPL WHERE ETB_CONGESPAYES="-"') then
      Plus := ' AND PMA_MOTIFABSENCE <> "PRI"';
  end;

  if Plus = '' then
    Grille.ColFormats[0]   := 'CB=PGMOTIFABSENCELIB' { PT13 }
  else
    Grille.ColFormats[0]   := 'CB=PGMOTIFABSENCELIB|' + Plus;
  //FIN PT23
  if StMode='CP' then  Grille.FixedCols := 1;
  Grille.ColTypes[1]     := 'D';
  Grille.ColFormats [1]  := ShortDateFormat;
  Grille.ColFormats[2]   := 'CB=PGDEMIJOURNEE';
  Grille.ColTypes[3]     := 'D';
  Grille.ColFormats [3]  := ShortDateFormat;
  Grille.ColFormats[4]   := 'CB=PGDEMIJOURNEE';
  Grille.ColTypes[5]     := 'D';
  Grille.ColFormats[5]  := ShortDateFormat;
  Grille.ColFormats[6]   := 'CB=PGOUINON';
  Grille.ColTypes[7]     := 'R';
  Grille.ColAligns[7]    := taRightJustify ;
  Grille.ColTypes[8]     := 'R';
  Grille.ColAligns[8]    := taRightJustify ;
  Grille.ColLengths[9]    := 35;

  AfficheValeurDefaut(Row,Vide);
end;
{ FIN PT10 }

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Saisie groupée des absences, attribue les valeurs par 
Suite ........ : défaut de la saisie
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.AfficheValeurDefaut(Row: integer ; vide : boolean );
var
  Q : TQuery; //PT23
begin

  if vide then
    begin
    Grille.Cells[1,Row] := '';
    Grille.CellValues[2,Row] := 'MAT';   { PT12 }
    Grille.Cells[3,Row] := '';
    Grille.CellValues[4,Row] := 'PAM';   { PT12 }
    Grille.Cells[5,Row] := '';
    Grille.CellValues[6,Row] := 'NON';   { PT12 }
    Grille.Cells[9,Row] := '';
    if StMode='CP' then  Grille.CellValues[0,Row] := 'PRI'
    else                 Grille.CellValues[0,Row] := '';
    end
  else
    Begin
    //DEB PT23
    if Plus <> '' then
    begin
      Q := OpenSQL('SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_TYPEMOTIF = "ABS" ' + Plus + ' ORDER BY PMA_LIBELLE',True);
      if not Q.Eof then
        Grille.CellValues[0,Row] := Q.Findfield('PMA_MOTIFABSENCE').AsString
      else
        Grille.CellValues[0,Row] := '';
      Ferme(Q);
    end
    //FIN PT23
    else
      Grille.CellValues[0,Row] := 'PRI';
    Grille.CellValues[1,Row] := DateToStr(Date);
    Grille.CellValues[2,Row] := 'MAT';
    Grille.CellValues[3,Row] := DateToStr(Date);
    Grille.CellValues[4,Row] := 'PAM';
    Grille.CellValues[5,Row] := DateToStr(Date);
    Grille.CellValues[6,Row] := 'NON';
    if Plus <> '' then      //PT23
      if Grille.Cells[0,row] <> '' then
        Grille.CellValues[9,Row] := Copy(Grille.Cells[0,row],5,18)+' '+ FormatDateTime('dd/mm/yy', StrToDate(Grille.Cells[1,Row])) + ' au ' + FormatDateTime('dd/mm/yy', StrToDate(Grille.Cells[3,Row]))
      else
        Grille.CellValues[9,Row] := ''
    else
      Grille.CellValues[9,Row] := 'Congés payés ' + FormatDateTime('dd/mm/yy', Date) + ' au ' + FormatDateTime('dd/mm/yy', Date);
    end;
end;
{ FIN PT10 }
{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Suppression d'une ligne de la girlle
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.GrilleDeleteLigne(sender: TObject);
var ARow : integer;
begin
if Grille.RowCount = 2 then  AfficheValeurDefaut (1,True)
else
  if Grille.RowCount>2 then
  Begin
  ARow := Grille.Row;
  Grille.CacheEdit;
  Grille.SynEnabled := False;
  Grille.DeleteRow(Grille.Row);
  if Arow < 2 then Grille.Row := 1 else Grille.Row := ARow-1;
  Grille.MontreEdit;
  Grille.SynEnabled := True;
  Grille.Col := Grille.FixedCols;
  End;
end;
{ FIN PT10 }
{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Insertion d'une ligne dans la grille
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.GrilleInsereLigne(sender: TObject);
Var ARow, Acol : Integer;
begin
  ARow := Grille.Row;
  Acol := Grille.Col;
  if Acol <> Grille.FixedCols then Begin Grille.ValCombo.Hide; Grille.colcombo := -1; end;  
  Grille.CacheEdit;
  Grille.SynEnabled := False;
  Grille.InsertRow(Grille.Row);
  Grille.Row := ARow;
 // Grille.Col := Grille.FixedCols;
  Grille.MontreEdit;
  Grille.SynEnabled := True;

  AfficheValeurDefaut (Arow,True);

end;
{ FIN PT10 }

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Contrôle en entrée de cellule de la grille
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.GrilleCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
//Affichage de l'elipsis
If (Grille.col=1) or (Grille.col=3) or (Grille.col=5) then
  Begin
  Grille.ElipsisButton:=True;
  Grille.ValCombo.Hide;
  end
else
  Grille.ElipsisButton:=False;
//Raffraichissement du libellé
If (Grille.col=0) or (Grille.col=1) or (Grille.col=3) then
   StValue := Grille.Cells[Grille.Col,Grille.Row];
end;
{ FIN PT10 }

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Sortie de la celule de la grille
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.GrilleCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
Var i : integer;
begin
//Exit zone journée
If ((Acol=2) or (Acol=4)) and (IsValidDate(Grille.Cells[1,ARow])) AND (IsValidDate(Grille.Cells[3,ARow])) then
  Begin
  if StrToDate(Grille.Cells[1,ARow])=StrToDate(Grille.Cells[3,ARow]) then
      if (Grille.CellValues[2,ARow]='PAM') AND (Grille.CellValues[4,ARow]='MAT') then
      Begin
      PgiBox('Vous ne pouvez saisir une absence de l''après-midi au matin sur la même journée!',StCaption);
      Grille.CellValues[4,ARow]:='PAM';
      End;
  End;
//Contrôle cohérence si pls absences saisies
DateError := False;
if (Grille.RowCount>2) then
  Begin
  if (IsValidDate(Grille.Cells[1,ARow])) AND (IsValidDate(Grille.Cells[3,ARow]))
    AND (Grille.Cells[2,ARow]<>'') AND (Grille.Cells[4,ARow]<>'') then
       Begin
       For i := 1 to Grille.RowCount-1 do
         Begin
         if i=ARow then continue;
         if (IsValidDate(Grille.Cells[1,i])) AND (IsValidDate(Grille.Cells[3,i]))
             AND (Grille.Cells[2,i]<>'') AND (Grille.Cells[4,i]<>'') then
            Begin
            //Si date début saisie inférieur date de fin existant
            If (StrToDate(Grille.Cells[1,Arow])=StrToDate(Grille.Cells[3,i]))
            AND (StrToDate(Grille.Cells[3,Arow])=StrToDate(Grille.Cells[1,i]))
            AND (StrToDate(Grille.Cells[3,Arow])=StrToDate(Grille.Cells[1,Arow]))
            AND (Grille.CellValues[2,Arow]=Grille.CellValues[4,Arow])
            AND (Grille.CellValues[2,i]=Grille.CellValues[4,i])
            AND (Grille.CellValues[2,i]<>Grille.CellValues[2,Arow]) then
              Continue;
            //Si date début saisie inférieur date de fin existant
            If (StrToDate(Grille.Cells[1,Arow])<StrToDate(Grille.Cells[3,i]))
               AND (StrToDate(Grille.Cells[1,Arow])>StrToDate(Grille.Cells[1,i])) then
               Begin
               PGiBox('Cette absence est à cheval sur une absence saisie!',StCaption);
               DateError := True;
               Break;
               End
            else
              If (StrToDate(Grille.Cells[1,Arow])=StrToDate(Grille.Cells[3,i]))
                 AND (Grille.CellValues[2,Arow]='MAT') THEN
                   Begin
                   PGiBox('Cette absence est à cheval sur une absence saisie!',StCaption);
                   DateError := True;
                   Break;
                   End
              else
                If (StrToDate(Grille.Cells[1,Arow])=StrToDate(Grille.Cells[3,i]))
                  AND (Grille.CellValues[2,Arow]='PAM') AND (Grille.CellValues[4,i]='PAM') then
                   Begin
                   PGiBox('Cette absence est à cheval sur une absence saisie!',StCaption);
                   DateError := True;
                   Break;
                   End
                else
                //Si date fin saisie Supérieur date de début existant
                 If (StrToDate(Grille.Cells[3,Arow])>StrToDate(Grille.Cells[1,i]))
                    and (StrToDate(Grille.Cells[3,Arow])<StrToDate(Grille.Cells[3,i]))  then
                    Begin
                    PGiBox('Cette absence est à cheval sur une absence saisie!',StCaption);
                    DateError := True;
                    Break;
                    End
                 else
                   If (StrToDate(Grille.Cells[3,Arow])=StrToDate(Grille.Cells[1,i]))
                     AND (Grille.CellValues[4,Arow]='PAM') Then
                       Begin
                       PGiBox('Cette absence est à cheval sur une absence saisie!',StCaption);
                       DateError := True;
                       Break;
                       End
                   else
                     If (StrToDate(Grille.Cells[3,Arow])=StrToDate(Grille.Cells[1,i]))
                       AND (Grille.CellValues[4,Arow]='MAT') AND (Grille.CellValues[2,i]='MAT')  then
                         Begin
                         PGiBox('Cette absence est à cheval sur une absence saisie!',StCaption);
                         DateError := True;
                         Break;
                         End;
            End;
         End;
       End;
    End;
SetControlEnabled('BINS_LINE',(Not DateError));
SetControlEnabled('BDEL_LINE',(Not DateError));
//Raffraichissement du libellé, réaffectation des dates
If ((Acol=0) or (Acol=1) or (Acol=3)) and (StValue<>Grille.Cells[ACol,ARow]) then
  Begin
  if (Acol=1)  and (not IsValidDate(Grille.Cells[3,ARow])) then Grille.Cells[3,ARow] := Grille.Cells[1,ARow];
  if ((Acol=1) or (Acol=3))  and (not IsValidDate(Grille.Cells[5,ARow])) then Grille.Cells[5,ARow] := Grille.Cells[3,ARow];
  if IsValidDate(Grille.Cells[1,ARow]) AND IsValidDate(Grille.Cells[3,ARow]) then
    Begin
    if (Acol=1) and (StrToDate(Grille.Cells[3,ARow])< StrToDate(Grille.Cells[1,ARow])) then
       Begin
       Grille.Cells[3,ARow] := Grille.Cells[1,ARow];
       Grille.Cells[5,ARow] := Grille.Cells[1,ARow];
       End;
    if (Acol=3) then
      if (StrToDate(Grille.Cells[3,ARow])< StrToDate(Grille.Cells[1,ARow])) then
       Begin
       PgiBox('La date de fin ne peut être antérieur à la date de début d''absence.',StCaption);
       Grille.Cells[3,ARow] := Grille.Cells[1,ARow];
       Grille.Cells[5,ARow] := Grille.Cells[1,ARow];
       End
    else
       Grille.Cells[5,ARow] := Grille.Cells[3,ARow];
    Grille.CellValues[9,ARow] := Copy(Grille.Cells[0,Arow],5,18)+' '+ FormatDateTime('dd/mm/yy', StrToDate(Grille.Cells[1,ARow])) + ' au ' + FormatDateTime('dd/mm/yy', StrToDate(Grille.Cells[3,ARow]));
    End;
  End;
//Supression du format date si aucune saisie
If ((Acol=5) or (Acol=1) or (Acol=3)) then
    Begin
    if (Grille.Cells[Acol,ARow])='  /  /    ' then  Grille.Cells[Acol,ARow]:='';
    if (Grille.Cells[Acol,ARow]<>'') then
       if Not IsValidDate(Grille.Cells[Acol,ARow]) then
         Begin
         PgiBox('La date saisie est incorrect : '+Grille.Cells[Acol,ARow], StCaption);
         Grille.Cells[Acol,ARow]:='';
         End;
    End;
//Exit zone calcul duree
if (Acol=6) then
  Begin
  Grille.ColEditables[7] := (Grille.CellValues[6,ARow]='NON');
  Grille.ColEditables[8] := (Grille.CellValues[6,ARow]='NON');
  if (Grille.CellValues[6,ARow]='OUI') Then
    Begin
    Grille.CellValues[7,ARow]:='';
    Grille.CellValues[8,ARow]:='';
    SaisieLigne:=True;
    End;
  End;
//Exit zone jours
if (Acol=7) or (Acol=8) then
  Begin
  if (Grille.CellValues[6,ARow]='NON') then
    Begin
    if (Grille.CellValues[ACol,ARow]='') then
      Begin
       PGIBox('Vous devez saisir une valeur!',StCaption);
       SaisieLigne:=False;
      End
    else
      if StrToFloat(Grille.CellValues[Acol,ARow])<0 then
       Begin
       PGIBox('Vous devez saisir une valeur positif!',StCaption);
       Grille.CellValues[ACol,ARow]:='';
       SaisieLigne:=False;
       End
    else
       SaisieLigne:=True;
    End
  else
  if (Grille.CellValues[6,ARow]='OUI') then
       SaisieLigne:=True;
  End;
end;
{ FIN PT10 }

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Evenement sur touche clavier
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.GrilleCellKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var K : Char;
begin
If (Grille.col=1) or (Grille.col=3) or (Grille.col=5) then
   Case Key of
     VK_SPACE : BEGIN
                K :='*';
                ParamDate (Ecran, Sender,K);
                END ;

   End;
end;
{ FIN PT10 }

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Double Click grille
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.GrilleDblClick(sender: TObject);
Var K : Char;
begin
If (Grille.col=1) or (Grille.col=3) or (Grille.col=5) then
  Begin
  K :='*';
  ParamDate (Ecran, Sender,K);
  End;
end;
{ FIN PT10 }


{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Click sur elipsis de la celule de la grille
Mots clefs ... : PAIE;CP,ABSENCE
*****************************************************************}
{ DEB PT10 }
procedure TOF_CongesGrp.GrilleElipsisClick(Sender: TObject);
var key : char;
begin
key := '*';
ParamDate (Ecran, Sender, Key);
end;
{ FIN PT10 }

Initialization
registerclasses([TOF_CongesGrp]) ;

end.



