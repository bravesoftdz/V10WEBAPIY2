{***********UNITE*************************************************
Auteur  ...... : MF
Créé le ...... : 05/07/2001
Modifié le ... :   /  /
Description .. :  Saisie de la fiche ORGANISME : Création, Modifi
                  -cation, Suppression
Mots clefs ... : PAIE, PGDUCS, PGORGANISMES
*****************************************************************}
{
PT1    : 22/11/2001 SB V563 Fiche de bug n°368 : Suppression d'un organisme
PT2    : 17/05/2002 PH V582 Acces annuaire DP Social
PT3    : 27/05/2002 MF V582 Recherche sur tablette PGINSTITUTION affinée en fct
                            du type d'organisme et de la nature DUCS.
                            Récupétartion SIRET 1er émetteur de EMETTEURSOCIAL
                            valeur par défaut pour les identifiants émetteur et
                            destinataire, et les qualifiant
                            Traitement lié au mode de paiement.
PT4    : 12/06/2002 VG V582 Si la nature DUCS est différente de "IRC", on force
                           le code institution à '' sur le changement de nature
                           ducs
PT6    : 02/07/2002 VG V585 Version S3
PT7    : 07/08/2002 MF V585
                            1- Ajout formatage par défaut du champ "ligne optique"
                            Répond au cas IRC pour lequel le formatage de ce
                            champ est fonction de la demandde de l'organisme.
                            2- Affichage de la banque en fonction du type de
                            paiement.
PT8    : 22/10/2002 MF V585 Rectification de la longueur d'un champ sur
                            la fiche organisme + onglet par défaut
PT9    : 14/01/2003 MF V591 Fiche qualite 10383 : Les codes organismes 991 à 999
                            (tablette PGTYPEORGANISME) sont réservés aux groupes
                            d'institutions IRC.
PT10   : 30/01/2003 MF V591
                            1-Traitement de la création d'un premier organisme sur
                            un établissement : Nouvelle procédure ChangeEtab
                            Résoud Pb création enregistrement avec clé à blanc
                            2-Modification du traitement de suppression d'un
                            organisme. Le contrôle se fait sur l'existence
                            de cotisation dans la table HISTOBULLETIN pour cet
                            organisme et cet établissement.
PT11   : 18/02/2003 SB V595 Inutile de tester l'existence d'un enr. avant de le supprimer
// **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
}
unit UTOMOrganismePaie;

interface
uses
{$IFDEF EAGLCLIENT}
UtileAGL,eFiche,eFichList,MaineAgl,
{$ELSE}
db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fiche,FichList,BanqueCp,FE_Main,
{$ENDIF}
      StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOM,UTOB,HTB97,
      PgOutils, PGIenv ;

Type
     TOM_OrganismePaie = Class(TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnDeleteRecord; override;
       private
// début PT3
       qualifiant : string;
       procedure EnterPaieMode(Sender: TObject);
       procedure ChangeEtab (Sender : TObject);    // PT10-1
// fin PT3
       procedure OnClickAnnuaire (Sender:TObject);
     END ;


implementation

uses P5Def ;
{ TOM_OrganismePaie }

procedure TOM_OrganismePaie.OnArgument(stArgument: String);
Var VC : THValComboBox ;
    BT : TToolbarButton97;
    Pages : TPageControl; //PT8

// début PT3
{$IFDEF EAGLCLIENT}
     PaieMode : THValComboBox;
     ChamplibG : THEdit;
{$ELSE}
     PaieMode : THDBValComboBox;
     ChamplibG : THDBEdit;
{$ENDIF}
// fin PT3
begin
inherited;
//PT6
{$IFDEF CCS3}
SetControlProperty('PDUCSEDI','TabVisible',False); // Onglet invisible
{$ENDIF}
//FIN PT6
VC:=THValcomboBox(GetControl('ETABLISSEMENT')) ;
if VC<>Nil then
   BEGIN
   VC.ItemIndex:=0 ;
   if Ecran is TFFicheListe then
      BEGIN
      TFFicheListe(Ecran).FRange:=VC.Value ;
{$IFNDEF EAGLCLIENT}
      TFFicheListe(Ecran).FListe.Columns[0].Visible:=FALSE ;
{$ELSE}
      TFFicheListe(Ecran).FListe.ColWidths[0] := 0 ;
{$ENDIF}
      VC.OnChange := ChangeEtab; // PT10-1
// début PT3
{$IFDEF EAGLCLIENT}
PaieMode := THValComboBox(GetControl ('POG_PAIEMODE'));
ChamplibG := THEdit(GetControl('POG_REGROUPEMENT'));
{$ELSE}
PaieMode := THDBValComboBox (GetControl ('POG_PAIEMODE'));
ChamplibG := THDBEdit(GetControl('POG_REGROUPEMENT'));
{$ENDIF}
if PaieMode<>nil then PaieMode.OnEnter:=EnterPaieMode;

if ChampLibG<>nil then ChampLibG.Plus:=' AND PIP_INSTITUTION LIKE "G%"';
// fin PT3

      END ;
   END ;

// PT2 17/05/2002 PH V582 Acces annuaire DP Social
BT := TToolbarButton97 (GetControl ('BRECUPANN'));
if BT <> NIL then BT.OnClick := OnClickAnnuaire;
// **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
if (PgRendModeFonc ()<> 'MULTI') AND (BT <> NIL) then SetControlVisible ('BRECUPANN', FALSE);
// d PT8
SetControlProperty('POG_LONGTOTALE','width',37);

Pages := TPageControl(GetControl('PAGES'));
if Pages<>nil then
   Pages.ActivePageIndex:=0;
// f PT8
   

end;

// Procedure à completer avec des tests de cohérence sur l'onglet affiliation
// Voir si on le fait sur Onchangefield ou sur OnUpdate pour des controles à postériori
procedure TOM_OrganismePaie.OnChangeField(F: TField);
var
{$IFDEF EAGLCLIENT}
NatureOrg : THValComboBox;
Champlib : THEdit;        //PT3
{$ELSE}
NatureOrg : THDBValComboBox;
Champlib : THDBEdit;        //PT3
{$ENDIF}
// début PT3
   Q : Tquery;
   SiretEmett : string;
   TypeOrganisme : string;
// fin PT3
   LigneOptique, Etab : string; // PT7-1
   i : integer; // PT7-1
begin
inherited;

{$IFDEF EAGLCLIENT}
NatureOrg := THValComboBox(getcontrol('POG_NATUREORG'));
ChampLib := THEdit (GetControl('POG_INSTITUTION'));     //PT3
{$ELSE}
NatureOrg := THDBValComboBox(getcontrol('POG_NATUREORG'));
ChampLib := THDBEdit (GetControl('POG_INSTITUTION'));     //PT3
{$ENDIF}

// début PT3
// Récup. Siret premier émetteur table EMETTEURSOCIAL
Q := OpenSQL ('SELECT PET_SIRET '+
              'FROM EMETTEURSOCIAL', True);
SiretEmett := '';  // PORTAGECWAS
if Not Q.EOF then
   begin
     SiretEmett := Q.FindField('PET_SIRET').AsString;
   end;
ferme (Q);
// fin PT3

if (F.FieldName ='POG_NATUREORG') or (F.FieldName ='POG_ORGANISME') then  // ??
      begin
// début PT3
      SetControlEnabled('POG_NOCONTEMET',False);
      SetControlEnabled('POG_INSTITUTION',False);
      if ChampLib<>nil then
        begin
          TypeOrganisme := GetField('POG_ORGANISME');
          if (TypeOrganisme = '003') then
              // AGIRC
              ChampLib.Plus:=' AND PIP_INSTITUTION LIKE "C%"'+
                             ' AND  PIP_INSTITUTION not LIKE "Z%"';
          if (TypeOrganisme = '004') then
              // ARRCO
              ChampLib.Plus:=' AND PIP_INSTITUTION LIKE "A%"'+
                             ' AND  PIP_INSTITUTION not LIKE "Z%"';
          if (TypeOrganisme <> '003') and
             (TypeOrganisme <> '004') and
             (TypeOrganisme <> '999') then
              // Autre que Agirc, Arrco, groupe
              ChampLib.Plus:=' AND PIP_INSTITUTION not LIKE "G%"'+
                             ' AND  PIP_INSTITUTION not LIKE "Z%"';
// PT9          if (TypeOrganisme = '999') then
          if (TypeOrganisme <> '') and
             (StrToInt(TypeOrganisme) >= 991) and
             (StrToInt(TypeOrganisme) <= 999) then
              // Groupe
              ChampLib.Plus:=' AND PIP_INSTITUTION  LIKE "G%"';
        end;
// fin PT3

      if (NatureOrg <> NIL)
         and (GetField('POG_LONGEDITABLE') = '0')
         and (GetField('POG_POSDEBUT') = '0')
         and (GetField('POG_LONGTOTAL') = '0')
         and (GetField('POG_POSTOTAL') = '0')
         then
         begin
// début PT3
         SetField ('POG_IDENTQUAL','5');
         SetField ('POG_IDENTEMET',SiretEmett);
         SetField ('POG_IDENTDEST',GetField('POG_SIRET'));
         qualifiant := '5';
// fin PT3

         if NatureOrg.value = '100' then         // URSSAF
            begin
            SetField ('POG_SOUSTOTDUCS','-');
            SetField ('POG_LONGEDITABLE','4');
            SetField ('POG_POSDEBUT','4');
            end;
         if NatureOrg.value = '200' then         // ASSEDIC
            begin
// début PT3
            SetField ('POG_IDENTQUAL','ZZZ');
            SetField ('POG_IDENTEMET','ES');
            SetField ('POG_IDENTDEST','H02S000DUCS');
            qualifiant := 'ZZZ';
// fin PT3

            SetField ('POG_SOUSTOTDUCS','X');
            SetField ('POG_LONGTOTAL','1');
            SetField ('POG_POSTOTAL','3');
            SetField ('POG_LONGEDITABLE','3');
            SetField ('POG_POSDEBUT','5');
            end;
         if NatureOrg.value = '300' then         // IRC
            begin
            SetField ('POG_SOUSTOTDUCS','X');
            SetField ('POG_LONGEDITABLE','5');
            SetField ('POG_POSDEBUT','3');
            end;
      if NatureOrg.value = '600' then            // MSA
            begin
            SetField ('POG_SOUSTOTDUCS','X');
            SetField ('POG_LONGTOTAL','2');
            SetField ('POG_POSTOTAL','3');
            SetField ('POG_LONGEDITABLE','5');
            SetField ('POG_POSDEBUT','3');
            end;
      if NatureOrg.value = '700' then            // BTP
            begin
            SetField ('POG_SOUSTOTDUCS','X');
            SetField ('POG_LONGTOTAL','2');
            SetField ('POG_POSTOTAL','3');
            SetField ('POG_LONGEDITABLE','5');
            SetField ('POG_POSDEBUT','3');
            end;
         end;
// début PT3
        if NatureOrg.value = '300' then         // IRC
          begin
            SetControlEnabled('POG_NOCONTEMET',True);
            SetControlEnabled('POG_INSTITUTION',True);
          end
        else
          begin
            SetControlEnabled('POG_NOCONTEMET',False);
            SetField ('POG_NOCONTEMET','');
            SetControlEnabled('POG_INSTITUTION',False);
            SetField ('POG_INSTITUTION','');           //PT4
          end;
// fin PT3
      end;

// début PT3
      if F.FieldName ='POG_SIRET' then
      // Contrôle du Siret
      begin
        if ((DS.State in [dsInsert]) and
           (GetField('POG_ORGANISME') <> '') and
           (GetField('POG_NATUREORG') <> '')) or
           (((DS.State in [dsBrowse]) or(DS.State in [dsEdit])) and
           (GetField('POG_SIRET') <> '')) then
          begin
           if (Length(GetField('POG_SIRET')) <> 14) and
              (Length(GetField('POG_SIRET')) <> 9) then
           begin
              PGIBox('! Attention, N° Siret incomplet','N° Siret');
              SetFocusControl('POG_SIRET');
           end
           else
            if ControlSiret(GetField('POG_SIRET')) <> True then
              begin
                 PGIBox('! Attention, N° Siret erroné','N° Siret');
                 SetFocusControl('POG_SIRET');
              end
            else
              begin
                 // alimentation identifiant destinataire
                 // Siret
                 if (GetField('POG_IDENTDEST')= '') and
                    (GetField('POG_IDENTQUAL')= '5') then
                     SetField ('POG_IDENTDEST',GetField('POG_SIRET'));
                 // Siren
                 if (GetField('POG_IDENTDEST')= '') and
                    (GetField('POG_IDENTQUAL')= '22') then
                     SetField ('POG_IDENTDEST',Copy(GetField('POG_SIRET'),1,9));
//                end;
              end;
          end;

      end;
      if F.FieldName ='POG_IDENTQUAL' then
        // traitement qualifiant d'identifiant
        begin
          if (GetField('POG_IDENTDEST') = '') then
            // qualifiant destinataire  non encore renseigné
            begin
              if (GetField('POG_IDENTQUAL') = '5') then
                 // Siret
                 SetField('POG_IDENTDEST',GetField('POG_SIRET'));
              if (GetField('POG_IDENTQUAL') = '22') then
                 // Siren
                 SetField('POG_IDENTDEST',Copy(GetField('POG_SIRET'),1,9));
              if (GetField('POG_IDENTQUAL') = 'ZZZ') and
                 (GetField('POG_NATUREORG')='200') then
                 // Assedic (par défaut H02S000DUCS si réel, H03S000DUCSsi test)
                 SetField('POG_IDENTDEST','H02S000DUCS');
            end
          else
            // qualifiant destinataire  déjà renseigné
            begin
              if (GetField('POG_IDENTQUAL') = '22') then
                // Siren
                begin
                 if (qualifiant = '5')  then
                    // (ancien qualifiant = siret)
                    SetField('POG_IDENTDEST',Copy(GetField('POG_IDENTDEST'),1,9));

                 if (qualifiant = 'ZZZ') then
                    // (ancien qualifiant = selon accord)
                    SetField('POG_IDENTDEST',Copy(GetField('POG_SIRET'),1,9));
                end;

              if (GetField('POG_IDENTQUAL') = '5') then
                 // Siret
                  SetField('POG_IDENTDEST',GetField('POG_SIRET'));

              if (GetField('POG_IDENTQUAL') = 'ZZZ') then
                // selon accord
                if (GetField('POG_NATUREORG')='200') then
                 // Assedic (par défaut H02S000DUCS si réel, H03S000DUCSsi test)
                 SetField('POG_IDENTDEST','H02S000DUCS')
                else
                 SetField('POG_IDENTDEST','');
            end;

          if (GetField('POG_IDENTEMET') = '') then
            // qualifiant émetteur non encore renseigné
            begin
              if (GetField('POG_IDENTQUAL') = '5') then
                  // Siret
                  SetField('POG_IDENTEMET', SiretEmett);
              if (GetField('POG_IDENTQUAL') = '22') then
                  //Siren
                  SetField('POG_IDENTEMET', Copy(SiretEmett,1,9));

              if (GetField('POG_IDENTQUAL') = 'ZZZ') then
                // selon accord
                if (GetField('POG_NATUREORG')='200') then
                  //  Assedic (par défaut ES....
                  SetField('POG_IDENTEMET', 'ES')
                else
                  SetField('POG_IDENTEMET', '');
            end
          else
            // qualifiant émetteur déjà renseigné
            begin
              if (GetField('POG_IDENTQUAL') = '5') then
                  // Siret
                  SetField('POG_IDENTEMET', SiretEmett);

              if (GetField('POG_IDENTQUAL') = '22') then
                begin
                 // Siren
                 if (qualifiant = '5') then
                    SetField('POG_IDENTEMET',Copy(GetField('POG_IDENTEMET'),1,9));
                 if (qualifiant = 'ZZZ') then
                    SetField('POG_IDENTEMET',Copy(SiretEmett,1,9));
                end;

              if (GetField('POG_IDENTQUAL') = 'ZZZ') then
                 if (GetField('POG_NATUREORG')='200') then
                    // Assedic (par défaut H02S000DUCS si réel, H03S000DUCSsi test)
                    SetField('POG_IDENTEMET','ES')
                 else
                    SetField('POG_IDENTEMET','');
            end;
        end;


      if F.FieldName ='POG_IDENTEMET' then
        // traitement identifiant émetteur
        begin
           if(GetField('POG_IDENTEMET') <> '') then
          begin
           if (GetField('POG_IDENTQUAL') = '5') or
              (GetField('POG_IDENTQUAL') = '22') then

           if (Length(GetField('POG_IDENTEMET')) <> 14) and
              (Length(GetField('POG_IDENTEMET')) <> 9) then
           begin
              PGIBox('! Attention, Siret ou Siren incomplet','Siret ou Siren');
              SetFocusControl('POG_IDENTEMET');
           end
           else
            if ControlSiret(GetField('POG_IDENTEMET')) <> True then
              begin
                 PGIBox('! Attention, Siret ou Siren erroné','Siret ou Siren');
                 SetFocusControl('POG_IDENTEMET');
              end;
            end;
        end;

        if F.FieldName ='POG_IDENTDEST' then
        // traitement identifiant destinataire
        begin
           if (GetField('POG_IDENTDEST') <> '') then
          begin
           if (GetField('POG_IDENTQUAL') = '5') or
              (GetField('POG_IDENTQUAL') = '22') then

           if (Length(GetField('POG_IDENTDEST')) <> 14) and
              (Length(GetField('POG_IDENTDEST')) <> 9) then
           begin
              PGIBox('! Attention, Siret ou Siren incomplet','Siret ou Siren');
              SetFocusControl('POG_IDENTDEST');
           end
           else
            if ControlSiret(GetField('POG_IDENTDEST')) <> True then
              begin
                 PGIBox('! Attention, Siret ou Siren erroné','Siret ou Siren');
                 SetFocusControl('POG_IDENTDEST');
              end;
          end;
        end;
      if F.FieldName ='POG_PAIEMODE' then
       begin
         SetField('POG_RIBDUCSEDI','');
         Etab := GetField('POG_ETABLISSEMENT');
        // si mode de paiement = Z10 alors il faut renseigner l'identification OPS
        if (GetField('POG_PAIEMODE') <> 'Z10') and
           (GetField('POG_PAIEMODE') <> '') then
         begin
          SetControlEnabled('POG_IDENTOPS',false);
          SetField('POG_IDENTOPS','');
         end
        else
         SetControlEnabled('POG_IDENTOPS',true);
// d PT7-2
        if (GetField('POG_PAIEMODE') = '30') or
           (GetField('POG_PAIEMODE') = '31') or
           (GetField('POG_PAIEMODE') = 'Z10') then
          // Virement, Prélèvement, Télérèglement
          begin
            SetControlEnabled('POG_RIBDUCSEDI', True);
            if (GetField('POG_PAIEMODE') = '30') then
              // Virement (banque d l'organisme)
              begin
{$IFDEF EAGLCLIENT}
                THValComboBox(GetControl ('POG_RIBDUCSEDI')).DataType := 'PGBQORG';
{$ELSE}
                THDBValComboBox (GetControl ('POG_RIBDUCSEDI')).DataType := 'PGBQORG';
{$ENDIF$}
              end
            else
              // Prélèvement ou télérèglement (banque réservée au paiement des charges sociales)
              begin
                if (GetField('POG_RIBDUCSEDI') = '') then
                  begin
                    Q := OpenSQL ('SELECT ETB_RIBDUCSEDI '+
                                  'FROM ETABCOMPL WHERE '+
                                  'ETB_ETABLISSEMENT = "'+Etab+'"', True);
                    if Not Q.EOF then
                      begin
                        SetField('POG_RIBDUCSEDI',Q.FindField('ETB_RIBDUCSEDI').AsString);
                      end;
                    ferme (Q);
                  end;
{$IFDEF EAGLCLIENT}
                THValComboBox(GetControl ('POG_RIBDUCSEDI')).DataType := 'TTBANQUECP';
{$ELSE}
                THDBValComboBox (GetControl ('POG_RIBDUCSEDI')).DataType := 'TTBANQUECP';
{$ENDIF$}
              end;
          end
        else
          SetControlEnabled('POG_RIBDUCSEDI', False);
       end;
// f PT7-2

// fin PT3
// d PT7-1
      if F.FieldName ='POG_NUMINTERNE' then
        begin
          // alimentation 30 1er caractères de la ligne optique
          if (GetField('POG_NUMINTERNE') <> '') AND
             (GetField('POG_LGOPTIQUE') = '') then
            begin
              if  (GetField('POG_NATUREORG') <> '200') then
                begin
                  LigneOptique := Copy(GetField('POG_NUMINTERNE'),1, length(GetField('POG_NUMINTERNE')));
                  if (length(GetField('POG_NUMINTERNE')) < 30) then
                    begin
                      for i := length(GetField('POG_NUMINTERNE'))+1 to 30 do
                       begin
                         LigneOptique := LigneOptique + '0';
                       end;
                    end;
                end
              else
               // ASSEDIC
               begin
                 LigneOptique := 'S2'+Copy(GetField('POG_NUMINTERNE'),1, length(GetField('POG_NUMINTERNE')));
                 if ((length(GetField('POG_NUMINTERNE'))+2) < 30) then
                   begin
                     for i := length(GetField('POG_NUMINTERNE'))+3 to 30 do
                       begin
                         LigneOptique := LigneOptique + '0';
                       end;
                   end;
               end;
             SetField('POG_LGOPTIQUE',LigneOptique);
            end;
        end;
// f PT7-1
end;

// PT2 17/05/2002 PH V582 Acces annuaire DP Social
procedure TOM_OrganismePaie.OnClickAnnuaire(Sender: TObject);
var retour,st  : String;
    Q          : TQuery ;
begin
if GetField ('POG_ORGANISME') = '' then exit;
Retour := AglLanceFiche ('PAY','PGANNUAIRE', '', '' , 'R');
if (Retour <> '') AND (retour <> 'VIDE') then
   begin
   St := 'SELECT * FROM ANNUAIRE WHERE ANN_CODEPER='+Retour;
   Q := OpenSql (st, TRUE);
   if NOT Q.EOF then
      begin
      DS.edit;
      SetField ('POG_LIBELLE',Q.FindField ('ANN_NOMPER').AsString);
      SetField ('POG_ADRESSE1', Q.FindField ('ANN_ALRUE1').AsString);
      SetField ('POG_ADRESSE2', Q.FindField ('ANN_ALRUE2').AsString);
      SetField ('POG_ADRESSE3', Q.FindField ('ANN_ALRUE3').AsString);
      SetField ('POG_CODEPOSTAL',Q.FindField ('ANN_ALCP').AsString);
      SetField ('POG_VILLE', Q.FindField ('ANN_ALVILLE').AsString);
      SetField ('POG_TELEPHONE',Q.FindField ('ANN_TEL1').AsString);
      SetField ('POG_FAX', Q.FindField ('ANN_FAX').AsString);
      SetField ('POG_EMAIL',Q.FindField ('ANN_EMAIL').AsString);
      SetField ('POG_SIRET',Q.FindField ('ANN_SIREN').AsString + Q.FindField ('ANN_CLESIRET').AsString);
      end;
   ferme (Q);
   end;
end;
// FIN PT2

procedure TOM_OrganismePaie.OnDeleteRecord ;
var ExisteCod   : Boolean ;
    NomChamp    :array[1..2] of string ;
    ValChamp    :array[1..2] of variant ;
    Q           : TQuery ;
    St          : string ;
    Nb          : integer ;
begin
// d PT10-2
  inherited;
  {Une rubrique de cotisation associée à cet organisme pour cet établissement
   a-t-elle déjà été utilisée ? }
  st := 'SELECT Count (*) NBRE FROM HISTOBULLETIN '+
        'WHERE PHB_ETABLISSEMENT = "' + GetField('POG_ETABLISSEMENT') + '" AND '+
        'PHB_ORGANISME = "'+ GetField('POG_ORGANISME') + '" AND '+
        'PHB_NATURERUB = "COT"' ;
  Q := OpenSql (st, TRUE) ;
  if NOT Q.EOF then
    nb := Q.FindField ('NBRE').AsInteger
  else
    nb := 0;
  Ferme (Q);

  if nb <> 0 then
  {suppression impossible}
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Certaines cotisations alimentent cet organisme,'+
                  '#13#10Vous ne pouvez le supprimer!';
  end
  else
  {suppression possible - La suppression de la ventilation associée est-elle possible?}
  {le même organisme existe-t'il pour un autre établissement?}
  begin
    st := 'SELECT Count (*) NBRE FROM ORGANISMEPAIE '+
        'WHERE POG_ETABLISSEMENT <> "' + GetField('POG_ETABLISSEMENT') + '" AND '+
        'POG_ORGANISME = "'+ GetField('POG_ORGANISME') + '"';
    Q := OpenSql (st, TRUE);
    if NOT Q.EOF then
      nb := Q.FindField ('NBRE').AsInteger
    else
      nb := 0;
    Ferme (Q);

// f PT10-2
    { La suppression de la ventilation est possible}
{ PT11 Mise en commentaire : inutile
    begin
      NomChamp[1] := 'PVO_TYPORGANISME';
      ValChamp[1]:=GetField('POG_ORGANISME');
      ExisteCod:=RechEnrAssocier('VENTIORGPAIE',NomChamp,ValChamp);
      if ExisteCod=TRUE then     }
  if (nb = 0) then
    ExecuteSQL('DELETE FROM VENTIORGPAIE WHERE ##PVO_PREDEFINI## PVO_TYPORGANISME="'+GetField('POG_ORGANISME')+'" ') ;
  ChargementTablette(TFFicheListe(Ecran).TableName,'');
  end;
end;

procedure TOM_OrganismePaie.OnLoadRecord;
var
   LigneOptique : string; // PT7-1
   i : integer; // PT7-1
begin
inherited;
SetControlEnabled('POG_ORGANISME', (DS.State in [dsInsert]));

// début PT3
qualifiant :=  GetField('POG_IDENTQUAL'); //

        // si mode de paiement = Z10 alors il faut renseigner l'identification OPS
        if (GetField('POG_PAIEMODE') <> 'Z10') then
         SetControlEnabled('POG_IDENTOPS',false)
        else
         SetControlEnabled('POG_IDENTOPS',true);

        // si IRC alors il faut renseigner le n° de contrat de l'émetteur chez le destinataire
        if  (GetField('POG_NATUREORG') <> '300') then
         SetControlEnabled('POG_NOCONTEMET',False)
        else
         SetControlEnabled('POG_NOCONTEMET',True);

// d PT7-1
     // alimentation 30 1er caractères de la ligne optique
     if (GetField('POG_NUMINTERNE') <> '') AND
        (GetField('POG_LGOPTIQUE') = '') then
       begin
         if  (GetField('POG_NATUREORG') <> '200') then
          begin
           LigneOptique := Copy(GetField('POG_NUMINTERNE'),1, length(GetField('POG_NUMINTERNE')));
           if (length(GetField('POG_NUMINTERNE')) < 30) then
            begin
             for i := length(GetField('POG_NUMINTERNE'))+1 to 30 do
              begin
               LigneOptique := LigneOptique + '0';
              end;
            end;
          end
         else
          // ASSEDIC
          begin
            LigneOptique := 'S2'+Copy(GetField('POG_NUMINTERNE'),1, length(GetField('POG_NUMINTERNE')));
            if ((length(GetField('POG_NUMINTERNE'))+2) < 30) then
             begin
               for i := length(GetField('POG_NUMINTERNE'))+3 to 30 do
                 begin
                  LigneOptique := LigneOptique + '0';
                 end;
             end;
          end;
         SetControlText('POG_LGOPTIQUE',LigneOptique);
       end;
// f PT7-1

// fin PT3
end;

procedure TOM_OrganismePaie.OnNewRecord;
var ETAB : THValComboBox;
begin
  inherited;
  ETAB := THValComboBox (Ecran.findComponent ('ETABLISSEMENT'));
  if ETAB <> NIL then
     begin
     if (DS.State in [dsInsert]) then SetField ('POG_ETABLISSEMENT', ETAB.Value);
     end;
   SetField ('POG_BASETYPARR','P');
   SetField ('POG_MTTYPARR','P');
   SetField('POG_LONGTOTALE','7');
   SetField('POG_PERIODCALCUL','X');
   SetField('POG_AUTPERCALCUL','X');
   SetField('POG_PERIODICITDUCS','');
   SetField('POG_AUTREPERIODUCS','');
// début PT3
   SetField('POG_PAIEGROUPE','-');
   SetField('POG_PAIEMODE','');
   SetField('POG_IDENTOPS','');
   SetField('POG_IDENTQUAL','');
   SetField('POG_IDENTEMET','');
   SetField('POG_IDENTDEST','');
// fin PT3
end;

procedure TOM_OrganismePaie.OnUpdateRecord;
var
{$IFDEF EAGLCLIENT}
    ControlOrg : THValComboBox;
{$ELSE}
    ControlOrg : THDBValComboBox;
{$ENDIF}
    st,ValLibelle : String;
begin
  inherited;
st:='';
{$IFDEF EAGLCLIENT}
ControlOrg:=THValComboBox(GetControl('POG_ORGANISME'));
{$ELSE}
ControlOrg:=THDBValComboBox(GetControl('POG_ORGANISME'));
{$ENDIF}
if ControlOrg <> NIL then
  begin
  ValLibelle:= RechDom('PGTYPEORGANISME',ControlOrg.Value,FALSE) ;
  if (ValLibelle='') OR (ValLibelle='Error') then LastError:=1;
  if LastError <> 0 then
   begin
   SetFocusControl('POG_ORGANISME');
   LastErrorMsg:='Le code organisme est obligatoirement renseigné';
   end;
  end;
st := GetField ('POG_ETABLISSEMENT');
if st = '' then
   begin
   SetFocusControl('POG_ETABLISSEMENT');
   LastErrorMsg:='Le code POG_ETABLISSEMENT est obligatoirement renseigné';
   end;
// début PT3
if (GetField ('POG_PAIEGROUPE')='X') and (GetField('POG_DUCSDOSSIER')<>'X') then
   begin
   SetFocusControl('POG_PAIEGROUPE');
   LastErrorMsg:='Ce n''est pas une ducs dossier, paiement groupé impossible';
   LastError:=1
   end;
// fin PT3

   //Rechargement des tablettes
If (LastError=0) and (Getfield('POG_ORGANISME')<>'') and (Getfield('POG_LIBELLE')<>'') then
begin
  ChargementTablette(TFFicheListe(Ecran).TableName,'');
  SetControlEnabled ('ETABLISSEMENT', TRUE);  // PT10-1
end;
end;

// début PT3
procedure TOM_OrganismePaie.EnterPaieMode (Sender: TObject);
begin
     SetControlEnabled('POG_IDENTOPS',True);
end;
// fin PT3

// début PT10-1
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/01/2003
Modifié le ... :   /  /    
Description .. : gestion du changement d'établissement
Suite ........ : Dans le cas où il n'existe pas déjà d'organisme sur cet 
Suite ........ : établissement on force la création d'un nouvel 
Suite ........ : enregistrement.
Mots clefs ... : PAIE, PGDUCS, PGORGANISME
*****************************************************************}
procedure TOM_OrganismePaie.ChangeEtab(Sender: TObject);
var VC  : THValComboBox ;
    st  : String;
    Q   : TQuery;
    Nb  : Integer;
begin
  VC:=THValcomboBox(GetControl('ETABLISSEMENT')) ;
  if VC<>Nil then
  begin
    st := 'SELECT COUNT (*) NBRE FROM ORGANISMEPAIE WHERE POG_ETABLISSEMENT ="'+VC.Value+'"' ;
    Q := OpenSql (st, TRUE) ;
    if NOT Q.EOF then
      Nb := Q.FindField ('NBRE').AsInteger
    else
      nb := 0 ;
      Ferme(Q) ;

    if Nb = 0 then
    begin
      DS.Insert ;
      SetField ('POG_ETABLISSEMENT', VC.Value);
      SetControlEnabled ('POG_ORGANISME', TRUE);
      SetControlEnabled ('ETABLISSEMENT', FALSE);
      OnNewRecord ;
    end;
  end;
end;
// fin PT10-1

Initialization
//mvi registerclasses([TOM_OrganismePaie,TOM_CUMULPAIE]) ;
registerclasses([TOM_OrganismePaie]) ;
end.
