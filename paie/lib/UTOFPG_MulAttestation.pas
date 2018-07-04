{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Gestion du multi critere gestion  des attestations ASSEDIC
Mots clefs ... : PAIE;ATTESTATION
*****************************************************************}
{
PT1   : 27/09/2002 VG V585 Ajout d'une case à cocher pour afficher tous les
                           salariés (sauf DUE) : Fiche qualité N°10239
PT2   : 22/09/2003 SB V_42 Intégration compatibilité CWAS pour appel des .pdf
PT3   : 26/09/2005 SB V_65 FQ 12349 Suppression contrôle existance des fichiers .pdf
PT4   : 29/03/2006 EPI V_65 Ajout gestion des processus (FQ N°12791)
               Passage numéro salarié en paramètre
PT5   : 19/04/2006 EPI V_65 Ajout arret du processus
                            Si salarié sélectionné appel direct édition DUE
PT6   : 11/12/2006 GGS V_80 FQ 12994
PT7   : 08/01/2007 GG  V_80 FQ 13452 pouvoir lancer l'édition de la DUE pour
                            plusieurs salariés à la fois
PT8   : 26/03/2007 RMA V_702 Ajout Attestation accident du travail MSA
                             & Attestation maladie MSA
PT9   : 03/05/2007 PH  V_70 FQ 14101 Bouton ouvrir invisible dans le cas Tous Salariés
PT10  : 17/07/2007 FC  V_72 FQ 14532 Situation par défaut à "Tous les salariés"
PT11  : 05/12/2007 RM  V_80 FQ14646 + FQ14102 + FQ14328 + FQ14324 DblClick + Bouvrir ne marche Pas
                            Revoir affichage des listes suivant le critére Situation
PT12  : 12/02/2008 RM  V_80 FQ15214 : correction sortie Pg qd création filtre
PT14  : 02/07/2008 JPP V_80 FQ 12994 Changement de requête SQL et utilisation du "Plus" dans la ComboBox
PT16  : 05/08/2008 NA  V_80 FQ 15610 PB oracle : convention en integer une zone à blanc
}

unit UTOFPG_MulAttestation;
interface

uses StdCtrls, Controls, Classes, sysutils,
{$IFNDEF EAGLCLIENT}
  db, HDB, Mul, Fe_Main,
{$ELSE}
  MaineAgl, eMul, windows,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, EntPaie, HTB97,
  Hqry, PgOutils, PGoutils2;

type
  TOF_PGMULATTES = class(TOF)
  private
    Fichier: string;
      // PT4 début appel processus
    AppelProc: boolean;
      // PT5 SelectSal : String;
      // PT4 fin
    Situation: THValComboBox;
    Q_Mul: THQuery; // Query pour changer la liste associee
    Mul_Titre, Typeattest, CpltWhere: string; //PT8
    procedure DblClickGrill(Sender: TObject);
    procedure Nouveau(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure ActiveWhere(Sender: TObject);
      // PT5 ajout fonction de sortie
    procedure ClickSortie(Sender: TObject);
      //PT6 attestation faites non faites
    procedure AFaireChange(Sender: Tobject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;


implementation

uses ed_tools, hstatus, Variants;

{ TOF_PGMULATTES }

procedure TOF_PGMULATTES.DblClickGrill(Sender: TObject);
var
//PT8 EditText, Fiche, StRepOrig, matricule: string;
  EditText, Fiche, matricule: string;
begin
// PORTAGECWAS
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
  if Fichier = 'ACCTRAVAIL' then Fiche := 'ACCIDENTTRAVAIL'
  else if Fichier = 'ACCTRAVAIL_MSA' then Fiche := 'ACCIDENTTRAV_MSA' //PT8
  else Fiche := Fichier;
  if fichier <> '' then
  begin
    EditText := GetControlText('OBJET');
    if EditText = Fichier then
    begin //B
(*  DEB PT2 Mise en commentaire, Utilisation de $STD au lieu de VH_Paie.PGCheminRech
if VH_Paie.PGCheminRech='' then
       Begin
       HShowmessage('2;Chemin de recherche!;Vous devez renseigner les chemins de recherche dans les paramètres société!;W;O;O;;;','','');
       Exit;
       End;
    if FileExists(VH_Paie.PGCheminRech+'\'+Fichier+'.pdf') then
       AglLanceFiche ('PAY',Fiche,'','',TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring+';MODIFICATION')
       else
       begin
       HShowmessage('2;Fichier Introuvable!;le fichier '+Fichier+'.pdf n''existe pas sous le répertoire spécifié!;W;O;O;;;','','');
       Exit;
       end;*)
//    StRepOrig:=ChangeStdDatPath('$STD\'+Fichier+'.pdf',True);
//     if FileExists(StRepOrig) then
//PT6
     // pt16 if StrToInt(GetControlText('SITUATION')) = 3 then begin
        if (getcontroltext('SITUATION') <> '') AND (StrToInt(GetControlText('SITUATION')) = 3 ) then begin  // pt16
         matricule := TFmul(Ecran).Q.FindField('PAS_SALARIE').asstring;
         if Trim(matricule) <> '' then //PT8
            AglLanceFiche('PAY', Fiche, '', '', matricule + ';MODIFICATION');
      //PT11 end;
      //PT11 if StrToInt(GetControlText('SITUATION')) = 2 then begin
      End    //PT11
      Else   //PT11
      Begin  //PT11
         matricule := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
         if Trim(matricule) <> '' then //PT8
            AglLanceFiche('PAY', Fiche, '', '', matricule + ';CREATION');
      end;
//Fin PT6
//     else
//       PgiBox('Abandon du traitement.Le fichier '+StRepOrig+' n''existe pas.',Ecran.caption);
    { FIN PT2 }
    end; //E
  end;
end;

procedure TOF_PGMULATTES.ExitEdit(Sender: TObject);
var edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGMULATTES.Nouveau(Sender: TObject);
var
{$IFDEF EAGLCLIENT}
  Liste: THGrid;
{$ELSE}
  Liste: THDBGrid;
{$ENDIF}
  EditText, Fiche: string;
  i: integer;
begin
  if Fichier = 'ACCTRAVAIL' then Fiche := 'ACCIDENTTRAVAIL'
  else if Fichier = 'ACCTRAVAIL_MSA' then Fiche := 'ACCIDENTTRAV_MSA' //PT8
  else Fiche := Fichier;
  if (fichier <> '') then
  begin
    EditText := GetControlText('OBJET');
    if EditText = Fichier then
    begin //B
      if Fichier = 'DUE' then
      begin
    //PT7
{$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
        Liste := THGrid(GetControl('FListe'));
{$ENDIF}
        if Liste <> nil then
        begin
          if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
          begin
            MessageAlerte('Aucun élément sélectionné');
            exit;
          end;
          if (Liste.AllSelected = TRUE) then
          begin
            InitMoveProgressForm(nil, 'Edition en cours', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
            InitMove(TFmul(Ecran).Q.RecordCount, '');
{$IFDEF EAGLCLIENT}
            if (TFMul(Ecran).bSelectAll.Down) then
              TFMul(Ecran).Fetchlestous;
{$ENDIF}
            TFmul(Ecran).Q.First;
            while not TFmul(Ecran).Q.EOF do
            begin
              MoveCur(False);
              AglLanceFiche('PAY', Fiche, '', '', TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring + ';CREATION');
              MoveCurProgressForm;
              TFmul(Ecran).Q.Next;
            end;
            Liste.AllSelected := False;
            TFMul(Ecran).bSelectAll.Down := Liste.AllSelected;
          end
          else
          begin
            InitMoveProgressForm(nil, 'Edition en cours', 'Veuillez patienter SVP ...', Liste.NbSelected, FALSE, TRUE);
            InitMove(Liste.NbSelected, '');
            for i := 0 to Liste.NbSelected - 1 do
            begin
              Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
              TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
              MoveCur(False);
              AglLanceFiche('PAY', Fiche, '', '', TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring + ';CREATION');
              MoveCurProgressForm;
            end;
            Liste.ClearSelected;
          end;
          FiniMove;
          FiniMoveProgressForm;
        end;
    //Fin PT7
      end
      else
        if TFmul(Ecran).Q.RecordCount > 0 then //PT8 on plante si la liste est vide
          AglLanceFiche('PAY', Fiche, '', '', TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring + ';CREATION');

    { PT3 Mise en commentaire
    if VH_Paie.PGCheminRech='' then
       Begin
       HShowmessage('2;Chemin de recherche!;Vous devez renseigner les chemins de recherche dans les paramètres société!;W;O;O;;;','','');
       Exit;
       End;
    if FileExists(VH_Paie.PGCheminRech+'\'+Fichier+'.pdf') then
       AglLanceFiche ('PAY',Fiche,'','',TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring+';CREATION')
       else
       begin
       HShowmessage('2;Fichier Introuvable!;le fichier '+Fichier+'.pdf n''existe pas sous le répertoire spécifié!;W;O;O;;;','','');
       Exit;
       end;  }
    end; //E
  end;
end;

procedure TOF_PGMULATTES.OnArgument(Arguments: string);
var
  Edit: THEdit;
{$IFNDEF EAGLCLIENT}
  FListe: THDBGrid;
{$ELSE}
  FListe: THGrid;
{$ENDIF}
// PT5 Btn : TToolBarButton97;
// PT5 ajout bouton STOP
  Btn, BTNSTOP: TToolbarButton97;
  Txt: string; // PT5
begin
  // PT4 début modif
  // PT4 Fichier:=Arguments;
  Mul_Titre := ''; //PT8
  Typeattest := 'MAL'; //PT8
  Txt := '';
  AppelProc := False;
  Fichier := Trim(ReadTokenPipe(Arguments, ';'));
  Txt := Trim(ReadTokenPipe(Arguments, ';'));
  if Txt = 'P' then
    AppelProc := True;
// PT4 fin

// PT4 if  Arguments='MALADIE' then
//PT8 Debut Ajout ===>
  CpltWhere := '';
  if Fichier = 'MALADIE_MSA' then
  begin
    Mul_Titre := 'Attestation maladie MSA : ';
    Typeattest := 'MSA';
    CpltWhere := ' And PAS_ASSEDICCAISSE = "MAL" '
  end;
  if Fichier = 'ACCTRAVAIL_MSA' then
  begin
    Mul_Titre := 'Attestation accident de travail MSA : ';
    Typeattest := 'MSA';
    CpltWhere := ' And PAS_ASSEDICCAISSE = "ACT" '
  end;
//PT8 Fin Ajout <====

  if Fichier = 'MALADIE' then
    Mul_Titre := 'Attestation maladie : '; //PT8
// PT8    TFMul(Ecran).Caption:='Attestation maladie';

// PT4 if  Arguments='DUE' then
  if Fichier = 'DUE' then
//PT1
  begin
    Mul_Titre := 'Déclaration unique d''embauche : '; //PT8
    //PT8 TFMul(Ecran).Caption:='Déclaration unique d''embauche';
    //PT7
    //PT11 SetControlVisible('ALLSAL', False);
    SetControlVisible('bSelectAll', True);
    SetControlProperty('FListe', 'MultiSelection', True);
    SetControlVisible ('BOUVRIR', FALSE);  //PT11
    //Fin PT7
  end;
//FIN PT1
// PT4 if  Arguments='ACCTRAVAIL' then
  if Fichier = 'ACCTRAVAIL' then
    Mul_Titre := 'Attestation accident de travail : '; //PT8
    //PT8 TFMul(Ecran).Caption:='Attestation accident de travail';

//PT8 UpdateCaption(TFMul(Ecran));

  Edit := ThEdit(getcontrol('PSA_SALARIE'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
{$IFNDEF EAGLCLIENT}
  FListe := THDBGrid(GetControl('FListe'));
{$ELSE}
  FListe := THGrid(GetControl('FListe'));
{$ENDIF}
  if FListe <> nil then Fliste.OnDblClick := DblClickGrill;

  Btn := TToolBarButton97(GetControl('BInsert'));
  if Btn <> nil then Btn.OnClick := Nouveau;

// PT5 début modif
  if AppelProc = True then
  begin
    SetControlVisible('BTNSTOP', TRUE);
    BTNSTOP := TToolbarButton97(GetControl('BTNSTOP'));
    if BTNSTOP <> nil then BTNSTOP.Onclick := ClickSortie;
  end;
// PT5 fin

  MajAttestations;
//PT6
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  Situation := THValComboBox(GetControl('SITUATION'));
  if Situation <> nil then
  Begin //PT11
    IF Fichier = 'DUE' Then SetControlProperty('SITUATION','DATATYPE','PGDUEMUL');  //PT11
    Situation.OnChange := AFaireChange;
  End;  //PT11
 // PT9
 //PT11 SetControlVisible ('BOUVRIR', FALSE);
  SetControlVisible ('TDATESELECT', FALSE); //PT11
  SetControlVisible ('DATESELECT', FALSE);  //PT11

  IF Fichier = 'DUE' Then
    SetControlText('SITUATION', '1')     //PT11
  else if Fichier = 'MALADIE' then
  begin
    SetControlProperty('SITUATION', 'Plus', 'AND CO_CODE<>"1"'); //PT14 on enlève le cas des 'Les salariés partis'
    SetControlText('SITUATION', '2')    //PT14  2 = Attestation à faire (par défaut)
  end
  Else
  begin
    SetControlText('SITUATION', '4');    //PT10
  end;
end;
//PT6
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gérard Guillaud-Saumur
Créé le ...... : 11/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGATTESTATION;PGMULSALARIE
*****************************************************************}

procedure TOF_PGMULATTES.AfaireChange(Sender: TObject);
var
  TT: TFMul;
  i: Integer;
begin
  If Fichier <> 'DUE' then  //  PT12
  begin
    if (GetControlText('SITUATION') <> '') then
      i := StrToInt(GetControlText('SITUATION'))
    else
    begin
      if AppelProc = False then
      begin
        i := 1;
        SetControlText('SITUATION', '4');
      end
      else
      begin
        i := 3;
        SetControlText('SITUATION', '3');
      end;
    end;
  end
  else      // PT12
    i := 1;

  //PT11 Debut ajout =====>
  If Fichier = 'DUE' then
  Begin
     case i of
       1: TFMul(Ecran).Caption := Mul_Titre + 'Liste des Salariés entrés';
       2: TFMul(Ecran).Caption := Mul_Titre + 'Liste des déclarations à faire';
       3: TFMul(Ecran).Caption := Mul_Titre + 'Liste des déclarations faites';
     end;
  End
  Else
  Begin
  //PT11 Fin ajout <=====
     case i of
     //PT8 1: TFMul(Ecran).Caption:= 'Liste des Salariés sortis';
     //PT8 2: TFMul(Ecran).Caption:= 'Liste des Attestations à préparer';
     //PT8 3: TFMul(Ecran).Caption:= 'Liste des Attestations effectuées';
     //PT8 4: TFMul(Ecran).Caption:= 'Liste des Salariés';
       1: TFMul(Ecran).Caption := Mul_Titre + 'Liste des Salariés sortis';
       2: TFMul(Ecran).Caption := Mul_Titre + 'Liste des Attestations à préparer';
       3: TFMul(Ecran).Caption := Mul_Titre + 'Liste des Attestations effectuées';
       4: TFMul(Ecran).Caption := Mul_Titre + 'Liste des Salariés';
     end;
  End; //PT11
  // PT9
  //PT11 Debut modif ==>
  //if i = 1 then SetControlVisible ('BOUVRIR', FALSE)
  //else SetControlVisible ('BOUVRIR', TRUE);
  If i = 1 then
  Begin
     SetControlVisible ('TDATESELECT', TRUE);
     SetControlVisible ('DATESELECT', TRUE);
     SetControlText('DATESELECT', '01/01/1900');
  End
  Else
  Begin
     SetControlVisible ('TDATESELECT', FALSE);
     SetControlVisible ('DATESELECT', FALSE);
  End;
  //PT11 Fin modif <==

  TT := TFMul(Ecran);
  if TT <> nil then
    UpdateCaption(TT);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/09/2002
Modifié le ... :   /  /
Description .. : Rechargement des paramètres
Mots clefs ... : PAIE;ATTESTATION
*****************************************************************}

procedure TOF_PGMULATTES.ActiveWhere(Sender: TObject);
begin
//PT6 mise en commentaire
//PT1
//if (GetCheckBoxState('ALLSAL') <> cbChecked) then
   // PT4 SetControlText('XX_WHERE', ' PSA_DATESORTIE="'+UsDateTime(iDate1900)+'" OR'+
   // PT4       '  PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+USDateTime(Date)+'"')
   // PT4 début modif
   // PT5 If AppelProc = False then
//	   SetControlText('XX_WHERE', ' PSA_DATESORTIE="'+UsDateTime(iDate1900)+'" OR'+
//  				        '  PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+USDateTime(Date)+'"')
   // PT5 else
   // PT5   begin
  // PT5    SetControlText('XX_WHERE', SelectSal);
   // PT5   end
   // PT4 fin

//else
//   SetControlText('XX_WHERE', ' PAS_TYPEATTEST = "MAL" AND'+
//                ' PSA_DATESORTIE >= "'+UsDateTime(StrToDate('01/01/'+Annee))+'" AND'+
//                ' PSA_DATESORTIE <= "'+UsDateTime(StrToDate('31/12/'+Annee))+'"';

   //FIN PT1
//end;

  //PT11 Debut ajout =====>
  If Fichier = 'DUE' then
  Begin
     If (GetControlText('SITUATION') = '1') then
     begin // liste des salaries Entrés
       if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGMULSALARIE');

         If (IsValidDate(GetControlText('DATESELECT'))) And (GetControlText('DATESELECT') <> '01/01/1900')
            Then SetControlText('XX_WHERE', ' PSA_DATEENTREE >= "'+UsDateTime(StrToDate(GetControlText('DATESELECT')))+'"')
            Else SetControlText('XX_WHERE', '');
     end;
  End
  Else
  Begin
     If (GetControlText('SITUATION') = '1') then
     begin // liste des salaries sorties
       if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGMULSALARIE');

         If (IsValidDate(GetControlText('DATESELECT'))) And (GetControlText('DATESELECT') <> '01/01/1900')
            Then SetControlText('XX_WHERE', ' PSA_DATESORTIE >= "'+UsDateTime(StrToDate(GetControlText('DATESELECT')))+'"')
            Else SetControlText('XX_WHERE', ' PSA_DATESORTIE <> "'+UsDateTime(StrToDate('01/01/1900'))+'"');
     end;
  //PT11 Fin ajout <============================
     if (GetControlText('SITUATION') = '2') then
     begin // liste des salaries n'ayant pas d'attestations maladie
       if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGMULSALARIE');
         //PT8 SetControlText('XX_WHERE', 'PSA_SALARIE NOT IN (SELECT PAS_SALARIE FROM ATTESTATIONS WHERE'+
         //PT8           ' PAS_TYPEATTEST = "MAL") AND PSA_SALARIE IN (SELECT PCN_SALARIE FROM ABSENCESALARIE'+
         //PT8           ' WHERE PCN_TYPECONGE = "MAL")');

       //DEB PT14
       if (Fichier = 'MALADIE_MSA') or (Fichier = 'ACCTRAVAIL_MSA') then
       begin
        SetControlText('XX_WHERE',
          'PSA_SALARIE NOT IN (SELECT PAS_SALARIE FROM ATTESTATIONS WHERE PAS_TYPEATTEST = "' + Typeattest + '"' + CpltWhere + ')'
          + ' AND PSA_SALARIE IN (SELECT PCN_SALARIE FROM ABSENCESALARIE WHERE PCN_TYPECONGE ="MAL"))');
       end
       else
       begin
        // Requete avant FQ12994
        //SetControlText('XX_WHERE',
        //  'PSA_SALARIE IN (SELECT PCN_SALARIE FROM ABSENCESALARIE WHERE PCN_TYPEMVT="ABS" AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE WHERE PMA_TYPEABS IN ("MAP","MAN") AND PMA_TYPEATTEST="'+Typeattest+'"))'+

        // FQ12994
        SetControlText('XX_WHERE', 'PSA_SALARIE IN ('+
        'SELECT PCN_SALARIE'+
        ' FROM ABSENCESALARIE'+
        ' LEFT JOIN ATTESTATIONS ON PAS_DATEARRET=PCN_DATEDEBUTABS AND PAS_REPRISEARRET>=PCN_DATEFINABS'+
        ' WHERE PAS_DATEARRET IS NULL AND PAS_REPRISEARRET IS NULL AND PCN_TYPEMVT="ABS" AND PCN_TYPECONGE IN ('+
        '  SELECT PMA_MOTIFABSENCE'+
        '  FROM MOTIFABSENCE'+
        '  WHERE PMA_TYPEABS IN ("MAP","MAN","MAT","PAT") AND PMA_TYPEATTEST="MAL") )');

        // Solution N°2 non testée
        {SetControlText('XX_WHERE',
          'PSA_SALARIE IN ('+
          ' SELECT PCN_SALARIE'+
          ' FROM ABSENCESALARIE'+
          ' WHERE PCN_TYPEMVT="ABS" AND PCN_TYPECONGE IN ('+
          '  SELECT PMA_MOTIFABSENCE'+
          '  FROM MOTIFABSENCE'+
          '  WHERE PMA_TYPEABS IN ("MAP","MAN","MAT","PAT") AND PMA_TYPEATTEST="MAL"'+
          ' )'+
          ' AND NOT EXISTS ('+
          '  SELECT 1'+
          '  FROM ATTESTATIONS'+
          '  WHERE PAS_SALARIE=PCN_SALARIE AND ((PAS_DATEARRET=PCN_DATEDEBUTABS AND PAS_REPRISEARRET>=PCN_DATEFINABS) OR (PAS_DATEARRET="" AND PAS_REPRISEARRET=""))'+
          ' )'+
          ')'+
          ' AND (PSA_DATESORTIE="" OR PSA_DATESORTIE<="'+USDateTime(V_PGI.DateEntree)+'")');}
       end;
       //FIN PT14
     end;
     if (GetControlText('SITUATION') = '3') then
     begin // liste des Salariés avec Attestations maladie
       if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGASSEDICFAITES');
                //PT8 SetControlText('XX_WHERE',' PAS_TYPEATTEST = "MAL"');
       SetControlText('XX_WHERE', ' PAS_TYPEATTEST = "' + Typeattest + '"' + CpltWhere);
     end;
     if (GetControlText('SITUATION') = '4') then
     begin // liste des salaries
       if Q_Mul <> nil then
         TFMul(Ecran).SetDBListe('PGMULSALARIE');
       SetControlText('XX_WHERE', '');
     end;
  End;  //PT11
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/09/2002
Modifié le ... :   /  /
Description .. : Reload de la fiche => rechargement des paramètres
Mots clefs ... : PAIE;ATTESTATION
*****************************************************************}

procedure TOF_PGMULATTES.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;

// PT5 gestion arret du processus

procedure TOF_PGMULATTES.ClickSortie(Sender: TObject);
var BTNAnn: TToolbarButton97;
begin
  BTNAnn := TToolbarButton97(GetControl('BAnnuler'));
  if BTNAnn <> nil then
  begin
    TFMul(Ecran).Retour := 'STOP';
    BTNAnn.Click;
  end;
end;

initialization
  registerclasses([TOF_PGMULATTES]);
end.

