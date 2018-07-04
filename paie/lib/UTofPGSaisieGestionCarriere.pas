{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 13/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSAISIEGESTIONCARRIERE ()
Mots clefs ... : TOF;PGSAISIEGESTIONCARRIERE
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
PT1   : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Caractéristique
                           activité" - FQ N°14568
}
Unit UTofPGSaisieGestionCarriere ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,uTob,HSysMenu,EntPaie,ParamSoc,HTB97,Vierge,P5DEF,PGOutilsFormation ;

Type
  TOF_PGSAISIEGESTIONCARRIERE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    LeSalarie : String;
    TobHistorique : Tob;
    Pages : TPageControl;
    Procedure AfficheInfosSal;
    procedure ChangeOnglet (Sender : TObject);
    procedure AfficheCompetences;
    Procedure AfficheFormations;
    Procedure AfficheHistorique;
    Procedure AfficheContrats;
    procedure AccesSaisieCompetences (Sender : TObject);
    procedure AccesContrattravail (Sender : TObject);
    procedure AccesFormations (Sender : TOBject);
    Procedure AjoutLigneHistorique (TobLigneHisto : Tob;Libelle,Valeur,Commentaire : String;DateModif : TDateTime);
    Procedure BCriteresHistoClick (Sender : TObject);
    procedure SuiviCarriereGrilles;
    end ;


Implementation

procedure TOF_PGSAISIEGESTIONCARRIERE.OnUpdate ;
Var Q : TQuery;
    TobSal,TobChamps : Tob;
    i : Integer;
    Champ : String;
begin
        Inherited;
        Q := OpenSQL('SELECT * FROM DECHAMPS where DH_PREFIXE="PSA"',True);
        TobChamps := Tob.Create('Leschamps',Nil,-1);
        TobChamps.LoadDetailDB('Leschamps','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM SALARIES WHERE PSA_SALARIE="'+LeSalarie+'"',True);
        TobSal := Tob.Create('SALARIES',Nil,-1);
        TobSal.LoadDetailDB('SALARIES','','',Q,False);
        Ferme(Q);
        If TobSal.Detail.Count <> 1 then Exit;
        For i := 0 to TobChamps.Detail.Count - 1 do
        begin
                Champ := TobChamps.Detail[i].GetValue('DH_NOMCHAMP');
                If GetControl(Champ) <> nil then
                begin
                        If TobChamps.Detail[i].GetValue('DH_TYPECHAMP') = 'DATE' then TobSal.Detail[0].PutValue(Champ,StrToDate(GetControlText(Champ)))
                        else If TobChamps.Detail[i].GetValue('DH_TYPECHAMP') = 'DOUBLE' then TobSal.Detail[0].PutValue(Champ,StrToFloat(GetControlText(Champ)))
                        else If TobChamps.Detail[i].GetValue('DH_TYPECHAMP') = 'INTEGER' then TobSal.Detail[0].PutValue(Champ,StrToInt(GetControlText(Champ)))
                        else TobSal.Detail[0].PutValue(Champ,GetControlText(Champ));
                end;
        end;
        TobSal.Detail[0].UpdateDB(False);
        TobSal.Free;
        TobChamps.Free;
end;

procedure TOF_PGSAISIEGESTIONCARRIERE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGSAISIEGESTIONCARRIERE.OnArgument (S : String ) ;
var Onglet : TTabSheet;
    Grille : THGrid;
    BHisto : TToolBarButton97;
    Combo : THMultiValComboBox;
    Num : Integer;
begin
  Inherited ;
        Pages := TPageControl(GetControl('PAGES'));
        Pages.ActivePage := TTabSheet(GetControl('PIDENTITE',true));
        LeSalarie := Trim(readTokenPipe(S,';'));
        Onglet := TTabSheet(GetControl('PFORMATIONS'));
        If Onglet <> Nil then Onglet.OnShow := ChangeOnglet;
        Onglet := TTabSheet(GetControl('PCOMPETENCE'));
        If Onglet <> Nil then Onglet.OnShow := ChangeOnglet;
        Onglet := TTabSheet(GetControl('PHISTO'));
        If Onglet <> Nil then Onglet.OnShow := ChangeOnglet;
        Onglet := TTabSheet(GetControl('PCONTRAT'));
        If Onglet <> Nil then Onglet.OnShow := ChangeOnglet;
        AfficheInfosSal;
        SuiviCarriereGrilles;
        Grille := THGrid(GetControl('GCOMPETENCES'));
        if Grille <> Nil then Grille.OnDblClick := AccesSaisieCompetences;
        Grille := THGrid(GetControl('GCOMPETEMPLOI'));
        if Grille <> Nil then Grille.OnDblClick := AccesSaisieCompetences;
        Grille := THGrid(GetControl('GCOMPETSAL'));
        if Grille <> Nil then Grille.OnDblClick := AccesSaisieCompetences;
        Grille := THGrid(GetControl('GCONTRATS'));
        if Grille <> Nil then Grille.OnDblClick := AccesContrattravail;
{        Grille := THGrid(GetControl('GINIT'));
        if Grille <> Nil then Grille.OnDblClick := AccesFormations;
        Grille := THGrid(GetControl('GEXTERNE'));
        if Grille <> Nil then Grille.OnDblClick := AccesFormations;
        Grille := THGrid(GetControl('GINTERNE'));
        if Grille <> Nil then Grille.OnDblClick := AccesFormations;}
        BHisto := TToolBarButton97(GetControl('BHISTO'));
        If BHisto <> Nil then BHisto.OnClick := BCriteresHistoClick;
        Combo := THMultiValComboBox(GetControl('CRITERESHISTO'));
        Combo.Items.Add('Code emploi');
        Combo.Values.Add('CODEEMPLOI');
        Combo.Items.Add('Libelle emploi');
        Combo.Values.Add('LIBELLEEMPLOI');
        Combo.Items.Add('Qualification');
        Combo.Values.Add('QUALIFICATION');
        Combo.Items.Add('Coefficient');
        Combo.Values.Add('COEFFICIENT');
        Combo.Items.Add('Indice');
        Combo.Values.Add('INDICE');
        If VH_Paie.PGLibCodeStat <> '' Then begin Combo.Items.Add(VH_Paie.PGLibCodeStat); Combo.Values.Add('CODESTAGE');end;
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                If Num=1 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat1); Combo.Values.Add('TRAVAILN1');end;
                If Num=2 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat2); Combo.Values.Add('TRAVAILN2');end;
                If Num=3 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat3); Combo.Values.Add('TRAVAILN3');end;
                If Num=4 Then begin Combo.Items.Add(VH_Paie.PGLibelleOrgStat4); Combo.Values.Add('TRAVAILN4');end;
        end;
        For Num := 1 to VH_Paie.PgNbSalLib do
        begin
                Combo.Items.Add(GetParamSoc('SO_PGSALLIB'+IntToStr(Num))+ ' mensuel');
                Combo.Values.Add('SALAIREMOIS'+ IntToStr(Num));
        end;
        For Num := 1 to VH_Paie.PgNbSalLib do
        begin
                Combo.Items.Add(GetParamSoc('SO_PGSALLIB'+IntToStr(Num))+ ' annuel');
                Combo.Values.Add('SALAIREANN'+ IntToStr(Num));
        end;
        Combo.Items.Add('Horaire hebdomadaire');
        Combo.Values.Add('PGHHORHEBDO');
        Combo.Items.Add('Horaire annuel');
        Combo.Values.Add('PGHHORANNUEL');
        Combo.Items.Add('Statut professionnel DADS');
        Combo.Values.Add('DADSPROF');
        Combo.Items.Add('Statut catégoriel DADS');
        Combo.Values.Add('DADSCAT');
{PT1
        Combo.Items.Add('Condition d''emploi');
}
        Combo.Items.Add('Caractéristique activité');
//FIN PT1        
        Combo.Values.Add('CONDEMPLOI');
        Ecran.Caption := 'Fiche carrière du salarié : '+LeSalarie+' '+RechDom('PGSALARIE',LeSalarie,False);
        UpdateCaption(TFVierge(Ecran));
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
end ;

Procedure TOF_PGSAISIEGESTIONCARRIERE.AfficheInfosSal;
Var Q : TQuery;
    TobSal,TobChamps : Tob;
    i : Integer;
    Champ : String;
begin
        Q := OpenSQL('SELECT * FROM DECHAMPS where DH_PREFIXE="PSA"',True);
        TobChamps := Tob.Create('Leschamps',Nil,-1);
        TobChamps.LoadDetailDB('Leschamps','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM SALARIES WHERE PSA_SALARIE="'+LeSalarie+'"',True);
        TobSal := Tob.Create('lessalries',Nil,-1);
        TobSal.LoadDetailDB('lessalries','','',Q,False);
        Ferme(Q);
        If TobSal.Detail.Count <> 1 then Exit;
        For i := 0 to TobChamps.Detail.Count - 1 do
        begin
                Champ := TobChamps.Detail[i].GetValue('DH_NOMCHAMP');
                If GetControl(Champ) <> nil then
                begin
                        If TobChamps.Detail[i].GetValue('DH_TYPECHAMP') = 'DATE' then SetControlText(Champ,DateToStr(TobSal.Detail[0].GetValue(Champ)))
                        else If TobChamps.Detail[i].GetValue('DH_TYPECHAMP') = 'DOUBLE' then SetControlText(Champ,FloatToStr(TobSal.Detail[0].GetValue(Champ)))
                        else If TobChamps.Detail[i].GetValue('DH_TYPECHAMP') = 'INTEGER' then SetControlText(Champ,IntToStr(TobSal.Detail[0].GetValue(Champ)))
                        else SetControlText(Champ,TobSal.Detail[0].GetValue(Champ));
                end;
        end;
        TobSal.Free;
        TobChamps.Free;
        If GetControlText('PSA_CODEEMPLOI') <> '' then SetControlCaption('TPSA_CODEEMPLOI',RechDom('PGCODEPCSESE',GetControlText('PSA_CODEEMPLOI'),False));
        If GetControlText('PSA_LIBELLEEMPLOI') <> '' then SetControlCaption('TPSA_LIBELLEEMPLOI',RechDom('PGLIBEMPLOI',GetControlText('PSA_LIBELLEEMPLOI'),False));
        If GetControlText('PSA_QUALIFICATION') <> '' then SetControlCaption('TPSA_QUALIFICATION',RechDom('PGLIBQUALIFICATION',GetControlText('PSA_QUALIFICATION'),False));
        If GetControlText('PSA_COEFFICIENT') <> '' then SetControlCaption('TPSA_COEFFICIENT',RechDom('PGLIBCOEFFICIENT',GetControlText('PSA_COEFFICIENT'),False));
        If GetControlText('PSA_INDICE') <> '' then SetControlCaption('TPSA_INDICE',RechDom('PGLIBINDICE',GetControlText('PSA_INDICE'),False));
        If GetControlText('PSA_NIVEAU') <> '' then SetControlCaption('TPSA_NIVEAU',RechDom('PGLIBNIVEAU',GetControlText('PSA_NIVEAU'),False));
//        If GetControlText('PSA_CONVENTION') <> '' then SetControlCaption('LBLPSA_CONVENTION',RechDom('',GetControlText('PSA_CONVENTION')));

{        SetControlText('PSA_SALARIE',Q.FindField('PSA_SALARIE').AsString);
        SetControlText('PSA_NUMEROSS',Q.FindField('PSA_NUMEROSS').AsString);
        SetControlText('PSA_LIBELLE',Q.FindField('PSA_LIBELLE').AsString);
        SetControlText('PSA_PRENOM',Q.FindField('PSA_PRENOM').AsString);
        SetControlText('PSA_NOMJF',Q.FindField('PSA_NOMJF').AsString);
        SetControlText('PSA_PRENOMBIS',Q.FindField('PSA_PRENOMBIS').AsString);
        SetControlText('PSA_SURNOM',Q.FindField('PSA_SURNOM').AsString);
        SetControlText('PSA_ETABLISSEMENT',Q.FindField('PSA_ETABLISSEMENT').AsString);
        SetControlText('PSA_ADRESSE1',Q.FindField('PSA_ADRESSE1').AsString);
        SetControlText('PSA_ADRESSE2',Q.FindField('PSA_ADRESSE2').AsString);
        SetControlText('PSA_ADRESSE3',Q.FindField('PSA_ADRESSE3').AsString);
        SetControlText('PSA_CODEPOSTAL',Q.FindField('PSA_CODEPOSTAL').AsString);
        SetControlText('PSA_VILLE',Q.FindField('PSA_VILLE').AsString);
        SetControlText('PSA_PAYS',Q.FindField('PSA_PAYS').AsString);
        SetControlText('PSA_TELEPHONE',Q.FindField('PSA_TELEPHONE').AsString);
        SetControlText('PSA_PORTABLE',Q.FindField('PSA_PORTABLE').AsString);
        SetControlText('PSA_SEXE',Q.FindField('PSA_SEXE').AsString);
        SetControlText('PSA_QUALIFICATION',Q.FindField('PSA_QUALIFICATION').AsString);
        SetControlText('PSA_LIBELLEEMPLOI',Q.FindField('PSA_LIBELLEEMPLOI').AsString);

        Ferme(Q);}

end;

procedure TOF_PGSAISIEGESTIONCARRIERE.SuiviCarriereGrilles;
var Grille : THGrid;
    i : Integer;
    Libelle : String;
begin
// Formations
Grille := THGrid(GetControl('GINIT'));
If Grille <> Nil then
begin
        Grille.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
        Grille.ColTypes[1] := 'D';
        Grille.colformats[1] :=  ShortDateFormat;
        Grille.ColFormats[2] := 'CB=PGCENTREFORMATION';
end;
Grille := THGrid(GetControl('GEXTERNE'));
If Grille <> Nil then
begin
        Grille.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
        Grille.ColTypes[1] := 'D';
        Grille.colformats[1] :=  ShortDateFormat;
        Grille.ColTypes[2] := 'D';
        Grille.colformats[2] :=  ShortDateFormat;
        Grille.ColFormats[3] := 'CB=PGCENTREFORMATION';
end;
Grille := THGrid(GetControl('GINTERNE'));
If Grille <> Nil then
begin
        Grille.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
        Grille.ColTypes[1] := 'D';
        Grille.colformats[1] :=  ShortDateFormat;
        Grille.ColTypes[2] := 'D';
        Grille.colformats[2] :=  ShortDateFormat;
        Grille.ColFormats[3] := 'CB=PGCENTREFORMATION';
end;
//Compétences
Grille := THGrid(GetControl('GCOMPETEMPLOI'));
If Grille <> Nil then
begin
                Grille.ColFormats[0] := 'CB=PGRHCOMPETENCES';
                Grille.ColFormats[1] := '# ##0.00';
                Grille.ColAligns[1] := taRightJustify;
                Grille.ColFormats[2] := '# ##0.00';
                Grille.ColAligns[2] := taRightJustify;
                Grille.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
                For i := 1 to 5 do
                begin
                        Libelle := RechDom('GCZONELIBRE','RH'+IntToStr(i),False);
                        If Libelle <> 'Table libre '+IntToStr(i) then
                        begin
                                Grille.ColCount := Grille.ColCount + 1;
                                Grille.ColFormats[Grille.ColCount-1] := 'CB=PGTABLELIBRERH'+IntToStr(i);
                                Grille.Titres.Add(Libelle);
                                Grille.UpdateTitres;
                        end;
                end;
end;
Grille := THGrid(GetControl('GCOMPETSAL'));
If Grille <> Nil then
begin
        Grille.ColFormats[0] := 'CB=PGRHCOMPETENCES';
        Grille.ColFormats[1] := '# ##0.00';
        Grille.ColAligns[1] := taRightJustify;
        For i := 1 to 5 do
        begin
                Libelle := RechDom('GCZONELIBRE','RH'+IntToStr(i),False);
                If Libelle <> 'Table libre '+IntToStr(i) then
                begin
                        Grille.ColCount := Grille.ColCount + 1;
                        Grille.ColFormats[Grille.ColCount-1] := 'CB=PGTABLELIBRERH'+IntToStr(i);
                        Grille.Titres.Add(Libelle);
                        Grille.UpdateTitres;
                end;
        end;
end;
//Historique
//Contrats
Grille := THGrid(GetControl('GCONTRATS'));
If Grille <> Nil then
begin
        Grille.ColFormats[0] := 'CB=PGTYPECONTRAT';
        Grille.ColTypes[1] := 'D';
        Grille.colformats[1] :=  ShortDateFormat;
        Grille.ColTypes[2] := 'D';
        Grille.colformats[2] :=  ShortDateFormat;
end;
end;

procedure TOF_PGSAISIEGESTIONCARRIERE.ChangeOnglet (Sender : TObject);
var Emploi : String;
    GHisto,GEmp,GSal,GContrat : THGrid;
    GrilleInit,GrilleInterne,GrilleExterne : THGrid;
    DateDeb,DateFin : TDateTime;
    THCriteres : THMultiValComboBox;
begin
        If TTabSheet(Sender).Name = 'PFORMATIONS' then   
        begin
             GrilleInit := THGrid(GetControl('GINIT'));
             GrilleInterne := THGrid(GetControl('GINTERNE'));
             GrilleExterne := THGrid(GetControl('GEXTERNE'));
             CarriereAfficheFormations(LeSalarie,GrilleInit,GrilleInterne,GrilleExterne);
             If (GetControl('HMTrad') <> Nil) And (GrilleInit    <> Nil) Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GrilleInit);
             If (GetControl('HMTrad') <> Nil) And (GrilleInterne <> Nil) Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GrilleInterne);
             If (GetControl('HMTrad') <> Nil) And (GrilleExterne <> Nil) Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GrilleExterne);
        end;
        If TTabSheet(Sender).Name = 'PCOMPETENCE' then
        begin
             Emploi := GetControlText('PSA_LIBELLEEMPLOI');
             GEmp := THGrid(GetControl('GCOMPETEMPLOI'));
             GSal := THGrid(GetControl('GCOMPETSAL'));
             CarriereAfficheCompetences(LeSalarie,Emploi,GEmp,GSal);
             If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GEmp);
             If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GSal);
        end;
        If TTabSheet(Sender).Name = 'PHISTO' then
        begin
             If IsValidDate(GetControlText('DATEHISTO')) then DateDeb := StrToDate(GetControlText('DATEHISTO'))
             else
             begin
                PGIBox(GetControlText('DATEHISTO')+' n''est pas une date correcte',Ecran.Caption);
                Exit;
             end;
             If IsValidDate(GetControlText('DATEHISTO1')) then DateFin := StrToDate(GetControlText('DATEHISTO1'))
             else
             begin
                  PGIBox(GetControlText('DATEHISTO1')+' n''est pas une date correcte',Ecran.Caption);
                  Exit;
             end;
             THCriteres := THMultiValComboBox(GetControl('CRITERESHISTO'));
             GHisto := THGrid(GetControl('GHISTO'));
             CarriereAfficheHisto(LeSalarie,DateDeb,DateFin,GHisto,THCriteres);
             If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GHisto);
        end;
        If TTabSheet(Sender).Name = 'PCONTRAT' then
        begin
             GContrat := THGrid(GetControl('GCONTRATS'));
             CarriereAfficheContrats(LeSalarie,GContrat);
             If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GContrat);
        end;
end;

Procedure TOF_PGSAISIEGESTIONCARRIERE.AfficheFormations;
Var Q : TQuery;
    TobFormationsInit,TobFormationsInterne,TobFormationsExt : Tob;
    HMTrad: THSystemMenu;
    GrilleInit,GrilleInterne,GrilleExterne : THGrid;
    i : Integer;
begin
        GrilleInit := THGrid(GetControl('GINIT'));
        For i := 1 to GrilleInit.RowCount -1 do
        begin
                GrilleInit.Rows[i].Clear;
        end;
        GrilleInterne := THGrid(GetControl('GINTERNE'));
        For i := 1 to GrilleInterne.RowCount -1 do
        begin
                GrilleInterne.Rows[i].Clear;
        end;
        GrilleExterne := THGrid(GetControl('GEXTERNE'));
        For i := 1 to GrilleExterne.RowCount -1 do
        begin
                GrilleExterne.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT PFO_CODESTAGE,PFO_DATEFIN,PFO_CENTREFORMGU FROM FORMATIONS WHERE PFO_SALARIE="'+LeSalarie+'"'+
        'AND PFO_NATUREFORM="004"',True);
        TobFormationsInit := Tob.Create('LesFormationsInit',Nil,-1);
        TobFormationsInit.LoadDetailDB('LesFormationsInit','','',Q,False);
        Ferme(Q);
        TobFormationsInit.PutGridDetail(THGrid(GetControl('GINIT')),False,True,'',False);
        if TobFormationsInit.Detail.count > 0 then GrilleInit.RowCount := TobFormationsInit.Detail.count + 1
        else GrilleInit.RowCount := 2;
        TobFormationsInit.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GINIT'))) ;
        Q := OpenSQL('SELECT PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN FROM FORMATIONS WHERE PFO_SALARIE="'+LeSalarie+'"'+
        'AND PFO_NATUREFORM="001"',True);
        TobFormationsInterne := Tob.Create('LesFormationsInit',Nil,-1);
        TobFormationsInterne.LoadDetailDB('LesFormationsInit','','',Q,False);
        Ferme(Q);
        TobFormationsInterne.PutGridDetail(THGrid(GetControl('GINTERNE')),False,True,'',False);
        if TobFormationsInterne.Detail.count > 0 then GrilleInterne.RowCount := TobFormationsInterne.Detail.count + 1
        else GrilleInterne.RowCount := 2;
        TobFormationsInterne.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GINTERNE'))) ;
        Q := OpenSQL('SELECT PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_CENTREFORMGU FROM FORMATIONS WHERE PFO_SALARIE="'+LeSalarie+'"'+
        'AND PFO_NATUREFORM="002"',True);
        TobFormationsExt := Tob.Create('LesFormationsInit',Nil,-1);
        TobFormationsExt.LoadDetailDB('LesFormationsInit','','',Q,False);
        Ferme(Q);
        TobFormationsExt.PutGridDetail(THGrid(GetControl('GEXTERNE')),False,True,'',False);
        if TobFormationsExt.Detail.count > 0 then GrilleExterne.RowCount := TobFormationsExt.Detail.count + 1
        else GrilleExterne.RowCount := 2;
        TobFormationsExt.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GEXTERNE'))) ;
end;

Procedure TOF_PGSAISIEGESTIONCARRIERE.AfficheCompetences;
var Salarie,Emploi,MessError,Competence,NiveauAtteint,Libelle : String;
    Q : TQuery;
    TobCompetencesSal,TS,TobCompetencesEmploi,TE,TobComparaison,TC : Tob;
    DegreMaitriseE,DegreMaitriseS : Double;
    a,i : Integer;
    GrilleEmploi,GrilleSalarie : THGrid;
    Salaire,Nb : Double;
    HMTrad: THSystemMenu;
begin
{        Grille := THGrid(GetControl('GCOMPETENCES'));
        For i := 1 to Grille.RowCount -1 do
        begin
                Grille.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT PCH_DATEVALIDATION,PCH_COMPETENCE,PCH_DEGREMAITRISE FROM RHCOMPETRESSOURCE '+
        'WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="'+LeSalarie+'"',True);
        TobCompetences := Tob.Create('LesCompetences',Nil,-1);
        TobCompetences.LoadDetailDB('LesCompetences','','',Q,False);
        Ferme(Q);
        TobCompetences.PutGridDetail(Grille,False,True,'',False);
        If TobCompetences.Detail.count > 0 then Grille.RowCount := TobCompetences.Detail.count + 1
        else Grille.RowCount := 2;
        TobCompetences.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPETENCES'))) ;}
        GrilleEmploi := THGrid(GetControl('GCOMPETEMPLOI'));
        GrilleSalarie := THGrid(GetControl('GCOMPETSAL'));
        Salarie := GetControlText('PSA_SALARIE');
        Emploi := GetControlText('PSA_LIBELLEEMPLOI');
        TobComparaison := Tob.Create('Comparatif',Nil,-1);
        Q := OpenSQL('SELECT * FROM STAGEOBJECTIF LEFT JOIN RHCOMPETENCES ON POS_COMPETENCE=PCO_COMPETENCE WHERE POS_NATOBJSTAGE="EMP" '+
        'AND POS_CODESTAGE="'+Emploi+'"',True);
        TobCompetencesEmploi := Tob.Create('emploi',Nil,-1);
        TobCompetencesEmploi.LoadDetailDB('emploi','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT PCH_COMPETENCE,PCH_DEGREMAITRISE,PCO_TABLELIBRERH1,PCO_TABLELIBRERH2,PCO_TABLELIBRERH3,PCO_TABLELIBRERH4'+
        ',PCO_TABLELIBRERH5 FROM RHCOMPETRESSOURCE LEFT JOIN RHCOMPETENCES ON PCH_COMPETENCE=PCO_COMPETENCE WHERE PCH_TYPERESSOURCE="SAL" '+
        'AND PCH_SALARIE="'+Salarie+'"',True);
        TobCompetencesSal := Tob.Create('Salarie',Nil,-1);
        TobCompetencesSal.LoadDetailDB('Salarie','','',Q,False);
        Ferme(Q);
        For i := 0 To TobCompetencesEmploi.Detail.Count -1 do
        begin
                Competence := TobCompetencesEmploi.Detail[i].GetValue('POS_COMPETENCE');
                DegreMaitriseE := TobCompetencesEmploi.Detail[i].GetValue('POS_DEGREMAITRISE');
                DegreMaitriseS := 0;
                TS := TobCompetencesSal.FindFirst(['PCH_COMPETENCE'],[Competence],False);
                While TS <> Nil do
                begin
                        If TS.GetValue('PCH_DEGREMAITRISE') > DegreMaitriseS then DegreMaitriseS := TS.GetValue('PCH_DEGREMAITRISE');
                        TS := TobCompetencesSal.FindNext(['PCH_COMPETENCE'],[Competence],False);
                end;
                If DegreMaitriseS >= DegreMaitriseE then NiveauAtteint := 'OUI'
                else NiveauAtteint := 'NON';
                TC := Tob.Create('FilleCompare',TobComparaison,-1);
                TC.AddChampSupValeur('COMPETENCE',Competence,False);
                TC.AddChampSupValeur('MAITRISEEMP',DegreMaitriseE,False);
                TC.AddChampSupValeur('MAITRISESAL',DegreMaitriseS,False);
                TC.AddChampSupValeur('ATTEINT',NiveauAtteint,False);
                For a := 1 to 5 do
                begin
                        Libelle := RechDom('GCZONELIBRE','RH'+IntToStr(a),False);
                        If Libelle <> 'Table libre '+IntToStr(a) then
                        begin
                                TC.AddChampSupValeur('LIBRE'+IntToStr(a),TobCompetencesEmploi.Detail[i].GetValue('PCO_TABLELIBRERH'+intToStr(a)),False);
                        end;
                end;
        end;
        TobComparaison.PutGridDetail(GrilleEmploi,False,True,'',False);
        if TobComparaison.Detail.Count > 0 then GrilleEmploi.RowCount := TobComparaison.Detail.Count + 1
        else GrilleEmploi.RowCount := 2;
        TobComparaison.Free;
        TobComparaison := Tob.Create('Comparatif',Nil,-1);
        For i := 0 to TobCompetencesSal.Detail.Count - 1 do
        begin
                Competence := TobCompetencesSal.Detail[i].GetValue('PCH_COMPETENCE');
                DegreMaitriseS := TobCompetencesSal.Detail[i].GetValue('PCH_DEGREMAITRISE');
                TE := TobCompetencesEmploi.FindFirst(['POS_COMPETENCE'],[Competence],False);
                If TE = Nil then
                begin
                        TC := Tob.Create('FilleCompare',TobComparaison,-1);
                        TC.AddChampSupValeur('COMPETENCE',Competence,False);
                        TC.AddChampSupValeur('MAITRISE',DegreMaitriseS,False);
                        For a := 1 to 5 do
                        begin
                                Libelle := RechDom('GCZONELIBRE','RH'+IntToStr(a),False);
                                If Libelle <> 'Table libre '+IntToStr(a) then
                                begin
                                        TC.AddChampSupValeur('LIBRE'+IntToStr(a),TobCompetencesSal.Detail[i].GetValue('PCO_TABLELIBRERH'+intToStr(a)),False);
                                end;
                        end;
                end;
        end;
        TobComparaison.PutGridDetail(GrilleSalarie,False,True,'',False);
        If TobComparaison.Detail.Count > 0 then GrilleSalarie.RowCount := TobComparaison.Detail.Count + 1
        else GrilleSalarie.RowCount := 2;
        TobCompetencesSal.Free;
        TobCompetencesEmploi.Free;
        TobComparaison.Free;
        HMTrad.ResizeGridColumns(GrilleSalarie) ;
        HMTrad.ResizeGridColumns(GrilleEmploi) ;
end;

Procedure TOF_PGSAISIEGESTIONCARRIERE.AfficheHistorique;
Var Q : TQuery;
    TobHisto,TMC,TobMesChamps : Tob;
    HMTrad: THSystemMenu;
    Grille : THGrid;
    i,x,Num : Integer;
    DateModif : TDateTime;
    Valeur,Commentaire,LaTablette,LeChamp,LeChampBool,LeType,LeLibelle : String;
    THCriteres : THMultiValComboBox;
    LesCriteres,ChampCritere : String;
    DateDeb,DateFin : TDatetime;
begin
        If IsValidDate(GetControlText('DATEHISTO')) then DateDeb := StrToDate(GetControlText('DATEHISTO'))
        else
        begin
                PGIBox(GetControlText('DATEHISTO')+' n''est pas une date correcte',Ecran.Caption);
                Exit;
        end;
        If IsValidDate(GetControlText('DATEHISTO1')) then DateFin := StrToDate(GetControlText('DATEHISTO1'))
        else
        begin
                PGIBox(GetControlText('DATEHISTO1')+' n''est pas une date correcte',Ecran.Caption);
                Exit;
        end;
        THCriteres := THMultiValComboBox(GetControl('CRITERESHISTO'));
        LesCriteres := THCriteres.Value;
        TobMesChamps := Tob.Create('LesCahmps',Nil,-1);
        //CodeEmploi
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_CODEEMPLOI',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BCODEEMPLOI',False);
        TMC.AddchampSupValeur('LATABLETTE','PGCODEPCSESE',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Code emploi',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //CodeEmploi
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_LIBELLEEMPLOI',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BLIBELLEEMPLOI',False);
        TMC.AddchampSupValeur('LATABLETTE','PGLIBEMPLOI',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Libelle emploi',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Qualification
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_QUALIFICATION',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BQUALIFICATION',False);
        TMC.AddchampSupValeur('LATABLETTE','PGLIBQUALIFICATION',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Qualification',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Coeff
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_COEFFICIENT',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BCOEFFICIENT',False);
        TMC.AddchampSupValeur('LATABLETTE','PGLIBCOEFFICIENT',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Coefficient',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Indice
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_INDICE',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BINDICE',False);
        TMC.AddchampSupValeur('LATABLETTE','PGLIBINDICE',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Indice',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Niveau
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_NIVEAU',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BNIVEAU',False);
        TMC.AddchampSupValeur('LATABLETTE','PGLIBNIVEAU',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Niveau',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Code Stat
        if VH_Paie.PGLibCodeStat <> '' then
        begin
                TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
                TMC.AddchampSupValeur('LECHAMP','PHS_CODESTAT',False);
                TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BCODESTAT',False);
                TMC.AddchampSupValeur('LATABLETTE','PGCODESTAT',False);
                TMC.AddchampSupValeur('LETYPE','S',False);
                TMC.AddchampSupValeur('LELIBELLE',VH_Paie.PGLibCodeStat,False);
                TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        end;
        //Champ libres
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
                TMC.AddchampSupValeur('LECHAMP','PHS_TRAVAILN'+IntToStr(Num),False);
                TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BTRAVAILN'+IntToStr(Num),False);
                TMC.AddchampSupValeur('LATABLETTE','PGTRAVAILN'+IntToStr(Num),False);
                TMC.AddchampSupValeur('LETYPE','S',False);
                TMC.AddchampSupValeur('LELIBELLE',GetParamSoc('SO_PGLIBORGSTAT'+IntToStr(Num)),False);
                TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        end;
        //Salaires
        For Num := 1 to VH_Paie.PgNbSalLib do
        begin
                TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
                TMC.AddchampSupValeur('LECHAMP','PHS_SALAIREMOIS'+IntToStr(Num),False);
                TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BSALAIREMOIS'+IntToStr(Num),False);
                TMC.AddchampSupValeur('LATABLETTE','',False);
                TMC.AddchampSupValeur('LETYPE','I',False);
                TMC.AddchampSupValeur('LELIBELLE',GetParamSoc('SO_PGSALLIB'+IntToStr(Num))+ ' mensuel',False);
                TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        end;
        For Num := 1 to VH_Paie.PgNbSalLib do
        begin
                TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
                TMC.AddchampSupValeur('LECHAMP','PHS_SALAIREANN'+IntToStr(Num),False);
                TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BSALAIREANN'+IntToStr(Num),False);
                TMC.AddchampSupValeur('LATABLETTE','',False);
                TMC.AddchampSupValeur('LETYPE','I',False);
                TMC.AddchampSupValeur('LELIBELLE',GetParamSoc('SO_PGSALLIB'+IntToStr(Num))+ ' annuel',False);
                TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        end;
        //Horaires
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_PGHHORHEBDO',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BPGHHORHEBDO',False);
        TMC.AddchampSupValeur('LATABLETTE','',False);
        TMC.AddchampSupValeur('LETYPE','I',False);
        TMC.AddchampSupValeur('LELIBELLE','Horaire hebdomadaire',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Horaires
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_PGHHORANNUEL',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BPGHHORANNUEL',False);
        TMC.AddchampSupValeur('LATABLETTE','',False);
        TMC.AddchampSupValeur('LETYPE','I',False);
        TMC.AddchampSupValeur('LELIBELLE','Horaire annuel',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Statut professionnel DADS
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_DADSPROF',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BDADSPROF',False);
        TMC.AddchampSupValeur('LATABLETTE','PGSPROFESSIONNEL',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Statut professionnel DADS',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Statut catégoriel DADS
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_DADSCAT',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BDADSCAT',False);
        TMC.AddchampSupValeur('LATABLETTE','PGSCATEGORIEL',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
        TMC.AddchampSupValeur('LELIBELLE','Statut catégoriel DADS',False);
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        //Condition d''emploi
        TMC := Tob.Create('TobMesCahmps',TobMesChamps,-1);
        TMC.AddchampSupValeur('LECHAMP','PHS_CONDEMPLOI',False);
        TMC.AddchampSupValeur('LECHAMPBOOL','PHS_BCONDEMPLOI',False);
        TMC.AddchampSupValeur('LATABLETTE','PGCONDEMPLOI',False);
        TMC.AddchampSupValeur('LETYPE','S',False);
{PT1
        TMC.AddchampSupValeur('LELIBELLE','Condition d''emploi',False);
}
        TMC.AddchampSupValeur('LELIBELLE','Caractéristique activité',False);
//FIN PT1
        TMC.AddchampSupValeur('PRISENCOMPTE','-',False);
        If LesCriteres <> '' then
        begin
                While LesCriteres <> '' do
                begin
                        ChampCritere := ReadTokenPipe(LesCriteres,';');
                        TMC := TobMesChamps.FindFirst(['LECHAMP'],['PHS_'+ChampCritere],False);
                        If TMC <> Nil Then TMC.PutValue('PRISENCOMPTE','X');
                end;
                TMC := TobMesChamps.FindFirst([''],[''],False);
                While TMC <> Nil do
                begin
                        If TMC.GetValue('PRISENCOMPTE') <> 'X' then TMC.Free;
                        TMC := TobMesChamps.FindNext([''],[''],False);
                end;
        end;
        Grille := THGrid(GetControl('GHISTO'));
        For i := 1 to Grille.RowCount -1 do
        begin
                Grille.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT * FROM HISTOSALARIE WHERE PHS_SALARIE="'+LeSalarie+'" '+
        'AND PHS_DATEEVENEMENT>="'+UsDateTime(DateDeb)+'" AND PHS_DATEEVENEMENT<="'+UsDateTime(DateFin)+'"',True);
        TobHisto := Tob.Create('Histo',Nil,-1);
        TobHisto.LoadDetailDB('Histo','','',Q,False);
        Ferme(Q);
        TobHistorique := Tob.Create('Edition',Nil,-1);
        For i := 0 to TobHisto.Detail.Count - 1 do
        begin
                DateModif := TobHisto.Detail[i].GetValue('PHS_DATEEVENEMENT');
                Commentaire := TobHisto.Detail[i].GetValue('PHS_COMMENTAIRE');
                If Copy(Commentaire,1,4) <> 'Init' then
                begin
                        For x := 0 To TobMesChamps.Detail.Count - 1 do
                        begin
                                LeChamp := TobMesChamps.Detail[X].GetValue('LECHAMP');
                                LeChampBool := TobMesChamps.Detail[X].GetValue('LECHAMPBOOL');
                                LaTablette := TobMesChamps.Detail[X].GetValue('LATABLETTE');
                                LeType := TobMesChamps.Detail[X].GetValue('LETYPE');
                                LeLibelle := TobMesChamps.Detail[X].GetValue('LELIBELLE');
                                If TobHisto.Detail[i].GetValue(LeChampBool) = 'X' then
                                    begin
                                        If LeType = 'S' then Valeur := TobHisto.Detail[i].Getvalue(LeChamp)
                                        else Valeur := FloatToStr(TobHisto.Detail[i].GetValue(LeChamp));
                                        If LaTablette <> '' then Valeur := RechDom(LaTablette,Valeur,False);
                                        AjoutLigneHistorique (TobHisto.Detail[i],LeLibelle,Valeur,Commentaire,DateModif);
                                end;
                        end;
                end;
        end;
        TobMesChamps.Free;
        TobHisto.Free;
        If TobHistorique.Detail.count > 0 then Grille.RowCount := TobHistorique.Detail.count + 1
        else Grille.RowCount := 2;
        TobHistorique.PutGridDetail(THGrid(GetControl('GHISTO')),False,True,'',False);
        TobHistorique.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GHISTO'))) ;
end;

Procedure TOF_PGSAISIEGESTIONCARRIERE.AfficheContrats;
Var Q : TQuery;
    TobContrats : Tob;
    HMTrad: THSystemMenu;
    Grille : THGrid;
    i : Integer;
begin
        Grille := THGrid(GetControl('GCONTRATS'));
        For i := 1 to Grille.RowCount -1 do
        begin
                Grille.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT FROM CONTRATTRAVAIL '+
        'WHERE PCI_SALARIE="'+LeSalarie+'"',True);
        TobContrats := Tob.Create('LesContrats',Nil,-1);
        TobContrats.LoadDetailDB('LesContrats','','',Q,False);
        Ferme(Q);
        TobContrats.PutGridDetail(THGrid(GetControl('GCONTRATS')),False,True,'',False);
        If TobContrats.Detail.count > 0 then Grille.RowCount := TobContrats.Detail.count + 1
        else Grille.RowCount := 2;
        TobContrats.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GCONTRATS'))) ;
end;

procedure TOF_PGSAISIEGESTIONCARRIERE.AccesSaisieCompetences (Sender : TObject);
begin
        AGLLanceFiche('PAY','SAL_COMPETENCES','','','GESTCARRIERESAL;'+GetControlText('PSA_SALARIE'));
        AfficheCompetences;
end;

procedure TOF_PGSAISIEGESTIONCARRIERE.AccesContrattravail (Sender : TObject);
begin
        AGLLanceFiche('PAY','MUL_CONTRAT','',LeSalarie,LeSalarie);
        AfficheContrats;
end;

procedure TOF_PGSAISIEGESTIONCARRIERE.AccesFormations (Sender : TObject);
begin
        If THGrid(Sender).Name = 'GINIT' then AGLLanceFiche('PAY','FORMATIONCARRIERE','','',LeSalarie+';INITIALE');
        If THGrid(Sender).Name = 'GINTERNE' then AGLLanceFiche('PAY','FORMATIONCARRIERE','','',LeSalarie+';INTERNE');
        If THGrid(Sender).Name = 'GEXTERNE' then AGLLanceFiche('PAY','FORMATIONCARRIERE','','',LeSalarie+';EXTERNE');
        AfficheFormations;
end;

Procedure TOF_PGSAISIEGESTIONCARRIERE.AjoutLigneHistorique (TobLigneHisto : Tob;Libelle,Valeur,Commentaire : String;DateModif : TDateTime);
var TE : Tob;
begin
        TE := Tob.Create('FilleEtat',TobHistorique,-1);
        TE.AddchampSupValeur('DATEMODIF',DateModif,False);
        TE.AddchampSupValeur('COMMENTAIRE',Commentaire,False);
        TE.AddchampSupValeur('LIBELLEMODIF',Libelle,False);
        TE.AddchampSupValeur('VALEURMODIF',Valeur,False);
end;

Procedure TOF_PGSAISIEGESTIONCARRIERE.BCriteresHistoClick (Sender : TObject);
var DateDeb,DateFin : TDateTime;
    GHisto : THGrid;
    THCriteres : THMultiValComboBox;
begin
        If IsValidDate(GetControlText('DATEHISTO')) then DateDeb := StrToDate(GetControlText('DATEHISTO'))
             else
             begin
                PGIBox(GetControlText('DATEHISTO')+' n''est pas une date correcte',Ecran.Caption);
                Exit;
             end;
             If IsValidDate(GetControlText('DATEHISTO1')) then DateFin := StrToDate(GetControlText('DATEHISTO1'))
             else
             begin
                  PGIBox(GetControlText('DATEHISTO1')+' n''est pas une date correcte',Ecran.Caption);
                  Exit;
             end;
             THCriteres := THMultiValComboBox(GetControl('CRITERESHISTO'));
             GHisto := THGrid(GetControl('GHISTO'));
             CarriereAfficheHisto(LeSalarie,DateDeb,DateFin,GHisto,THCriteres);
             If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(GHisto);
end;

Initialization
  registerclasses ( [ TOF_PGSAISIEGESTIONCARRIERE ] ) ;
end.

