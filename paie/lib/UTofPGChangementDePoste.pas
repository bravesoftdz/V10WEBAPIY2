{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 15/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCHANGEMENTPOSTE ()
Mots clefs ... : TOF;PGCHANGEMENTPOSTE
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
PT1 24/03/2006 V_65 JL FQ 12958 Gestion accès bouton imprimer pour ne pas éditer un état pratiquement vide
PT2 10/08/2007 V_80 JL FQ 12957 Clic sur compétence active formation correspondante
}
Unit UTofPGChangementDePoste ;

Interface

Uses StdCtrls,Controls,Classes,UtobDebug,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,EdtREtat,
{$else}
     eMul,MainEAGL,UtilEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,uTob,HTB97,Vierge,HSysMenu,Grids,ParamSoc ;

Type
  TOF_PGCHANGEMENTPOSTE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad               ; override ;
    private
    LeSalarie : String;
    BCherche : TToolBarButton97;
    GrilleComp,GrilleCompSal,GrilleRecherche,GrilleForm : THGrid;
    Pages : TPageControl;
    procedure RemplirLesGrilles;
    procedure GRechercheOnDblClick(Sender : TObject);
    procedure ChercherFormations;
    procedure AfficherSelection(Sender : TObject);
    procedure AfficherSelectionComp(Sender : TObject); //PT2
    procedure AppliquerCriteres (Sender : TObject);
    procedure MiseEnFormeGrille;
    procedure ChangeOnglet(Sender : TObject);
    procedure GrilleDblCilck (Sender : TObject);
    procedure AfficheLibEmploi (Sender : TObject);
    Procedure ZoomChoixCompet(Sender : TObject);
    procedure BImprimerClick(Sender : TOBject);
    procedure BExportClick(Sender : TObject);
    procedure ImpressionSimulation(Exporter : Boolean);
    procedure RetaillerGrille(Sender : TObject);
   {procedure GrilleMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure DeposeCompetSal;
    Procedure DeposeCompet;
    procedure TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);}
  end ;

Implementation

procedure TOF_PGCHANGEMENTPOSTE.OnLoad;
begin
Inherited ;
SetControlEnabled('BImprimer',False);   //PT1
end;

procedure TOF_PGCHANGEMENTPOSTE.OnArgument (S : String ) ;
var Q : TQuery;
    Nom,Prenom : String;
    Onglet : TTabSheet;
    Edit : THEdit;
    BZoom,BImp : TToolBarButton97;
    i : Integer;
begin
  Inherited ;
        TFVierge(Ecran).OnResize := RetaillerGrille;
        Pages := TPageControl(GetControl('PAGES'));
        SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
        LeSalarie := ReadTokenPipe(S,';');
        Pages.ActivePage := TTabSheet(GetControl('PPOSTE',true));
        Onglet := TTabSheet(GetControl('PCOMPETENCES'));
        GrilleComp := THGrid(GetControl('GCOMPETENCES'));
        GrilleCompSal := THGrid(GetControl('GCOMPSAL'));
        GrilleRecherche := THGrid(GetControl('GRECHERCHE'));
        GrilleForm := THGrid(GetControl('GFORMATIONS'));
        If Onglet <> Nil then Onglet.OnShow := ChangeOnglet;
        If LeSalarie <> '' then
        begin
                Q := OpenSQL('SELECT PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM,PSA_LIBELLEEMPLOI FROM SALARIES WHERE PSA_SALARIE="'+LeSalarie+'"',True);
                If Not Q.eof then
                begin
                        Nom := Q.FindField('PSA_LIBELLE').AsString;
                        PreNom := Q.FindField('PSA_PRENOM').AsString;
                        SetControlText('ETABLISSEMENT',Q.FindField('PSA_ETABLISSEMENT').AsString);
                        SetControlText('EMPLOIACT',Q.FindField('PSA_LIBELLEEMPLOI').AsString);
                        SetControlText('EMPLOIFUT',Q.FindField('PSA_LIBELLEEMPLOI').AsString);
                        Ecran.Caption := 'Simulation changement de poste pour le salarié : '+LeSalarie+' '+Nom+' '+Prenom;
                        UpdateCaption(TFVierge(Ecran));
                end;
                Ferme(Q);
        end;
        SetControlText ('SALARIE', LeSalarie);
        SetControlCaption ('LIBEMPLOIACT' , RechDom ('PGLIBEMPLOI', GetControlText ('EMPLOIACT'), FALSE ));
        SetControlCaption ('LIBEMPLOIFUT' , RechDom ('PGLIBEMPLOI', GetControlText ('EMPLOIACT'), FALSE ));
        MiseEnFormeGrille;
        Edit := THEdit(GetControl('EMPLOIFUT'));
        If Edit <> Nil then Edit.ONChange := AfficheLibEmploi;
        BCherche := TToolBarButton97(GetControl('BAPPLIQUER'));
        If BCherche <> Nil then BCherche.OnClick := AppliquerCriteres;
        RemplirLesGrilles;
        GrilleRecherche.OnDblClick := GRechercheOnDblClick;
        GrilleRecherche.RowCount := 2;
        Ecran.Caption := 'Simulation évolutions pour le salarié '+ LeSalarie + ' ' + RechDom('PGSALARIE',LeSalarie,False);
        UpdateCaption(Ecran);
//        GrilleComp.Options := GrilleComp.Options + [goRowSelect] ;
//        GrilleCompSal.Options := GrilleCompSal.Options + [goRowSelect] ;
        GrilleForm.Options := GrilleForm.Options + [goRowSelect] ;
//        GrilleRecherche.Options := GrilleRecherche.Options + [goRowSelect] ;
        GrilleForm.OnClick := AfficherSelection;
        GrilleRecherche.OnClick := AfficherSelectionComp;      //PT2
        GrilleComp.OnDblClick := GrilleDblCilck;
        GrilleCompSal.OnDblClick := GrilleDblCilck;
        {  GrilleComp.OnMouseDown := GrilleMouseDown;
  GrilleCompSal.OnMouseDown := GrilleMouseDown;
  grilleRecherche.OnDragDrop := DEPOSE_OBJET;
  grilleRecherche.OnDragOver := TEST_DEPOSE_OBJET;}
        AppliquerCriteres(BCherche);
        BZoom := TToolBarButton97(GetControl('BZOOMCOMP'));
        SetControlVisible('BZOOMSAL',False);
        If BZoom <> Nil then BZoom.OnClick := ZoomChoixCompet;
        BImp := TToolBarButton97(GetControl('BImprimer'));
        If BImp<> Nil then BImp.OnClick := BImprimerClick;
        BImp := TToolBarButton97(GetControl('BEXPORTER'));
        If BImp<> Nil then BImp.OnClick := BExportClick;
        For i := 1 to 3 do
        begin
                SetControlText('LIBRE'+IntToStr(i),RechDom('GCZONELIBRE','CP'+IntToStr(i),False));
        end;
        For i := 1 to 5 do
        begin
                SetControlText('LLIBRERH'+IntToStr(i),RechDom('GCZONELIBRE','RH'+IntToStr(i),False));
        end;
end ;

procedure TOF_PGCHANGEMENTPOSTE.AfficheLibEmploi (Sender : TObject);
begin
        If Length (GetControlText('EMPLOIFUT')) = 3 then SetControlCaption ('LIBEMPLOIFUT' , RechDom ('PGLIBEMPLOI', GetControlText ('EMPLOIFUT'), FALSE ))
        else SetControlCaption('LIBEMPLOIFUT','');
end;

procedure TOF_PGCHANGEMENTPOSTE.AppliquerCriteres (Sender : TObject);
var Salarie,Emploi,MessError,Competence,NiveauAtteint : String;
    Q : TQuery;
    TobCompetencesSal,TS,TobCompetencesEmploi,TE,TobComparaison,TC : Tob;
    DegreMaitriseE,DegreMaitriseS : Double;
    i,a : Integer;
    GrilleEmploi,GrilleSalarie : THGrid;
    Salaire,Nb : Double;
    HMTrad: THSystemMenu;
    Libelle : String;
begin
     SetControlEnabled('BImprimer',True);//PT1
        If Pages.ActivePage.PageIndex = 0 then
        begin
                Q := OpenSQL('SELECT SUM(PHC_MONTANT) MONTANT FROM PAIEENCOURS '+
                'LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="01" '+
                 'WHERE PPU_LIBELLEEMPLOI="'+GetControlText('EMPLOIACT')+'"',True);
                If Not Q.Eof then Salaire := Q.FindField('MONTANT').AsFloat
                else Salaire := 0;
                Ferme(Q);
                Q := OpenSQL('SELECT COUNT(PPU_SALARIE) NB FROM PAIEENCOURS WHERE PPU_LIBELLEEMPLOI="'+GetControlText('EMPLOIACT')+'"',True);
                If Not Q.Eof then Nb := Q.FindField('NB').AsFloat
                else Nb := 0;
                Ferme(Q);
                If Nb <> 0 then Salaire := Salaire / Nb
                else Salaire := 0;
                SetControlText('SALAIREMOYENACT',FloatToStr(Salaire));
                Q := OpenSQL('SELECT SUM(PHC_MONTANT) MONTANT FROM PAIEENCOURS '+
                'LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="01" '+
                 'WHERE PPU_LIBELLEEMPLOI="'+GetControlText('EMPLOIFUT')+'"',True);
                If Not Q.Eof then Salaire := Q.FindField('MONTANT').AsFloat
                else Salaire := 0;
                Ferme(Q);
                Q := OpenSQL('SELECT COUNT(PPU_SALARIE) NB FROM PAIEENCOURS WHERE PPU_LIBELLEEMPLOI="'+GetControlText('EMPLOIFUT')+'"',True);
                If Not Q.Eof then Nb := Q.FindField('NB').AsFloat
                else Nb := 0;
                Ferme(Q);
                If Nb <> 0 then Salaire := Salaire / Nb
                else Salaire := 0;
                SetControlText('SALAIREMOYENFUT',FloatToStr(Salaire));
                Q := OpenSQL('SELECT SUM(PHC_MONTANT) MONTANT FROM HISTOCUMSAL WHERE PHC_CUMULPAIE="01" AND PHC_SALARIE="'+LeSalarie+'"',True);
                If Not Q.Eof then Salaire := Q.FindField('MONTANT').AsFloat
                else Salaire := 0;
                Ferme(Q);
                Q := OpenSQL('SELECT COUNT(PPU_SALARIE) NB FROM PAIEENCOURS WHERE PPU_SALARIE="'+LeSalarie+'"',True);
                If Not Q.Eof then Nb := Q.FindField('NB').AsFloat
                else Nb := 0;
                Ferme(Q);
                If Nb <> 0 then Salaire := Salaire / Nb
                else Salaire := 0;
                SetControlText('SALAIREACT',FloatToStr(Salaire));
                GrilleEmploi := THGrid(GetControl('GEMPLOI'));
                GrilleSalarie := THGrid(GetControl('GSALARIE'));
                Salarie := GetControlText('SALARIE');
                Emploi := GetControlText('EMPLOIFUT');
                MessError := '';
                If Salarie = '' then MessError := '- Le matricule salarié';
                If Emploi= '' then MessError := '- L''emploi futur';
                If MessError <> '' then
                begin
                        PGIBox('Recherche impossible, vous devez renseigner : ',Ecran.Caption);
                        Exit;
                end;
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
                HMTrad.ResizeGridColumns(GrilleEmploi) ;
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
                HMTrad.ResizeGridColumns(GrilleSalarie) ;
                If TobComparaison.Detail.Count > 0 then GrilleSalarie.RowCount := TobComparaison.Detail.Count + 1
                else GrilleSalarie.RowCount := 2;
                TobCompetencesSal.Free;
                TobCompetencesEmploi.Free;
                TobComparaison.Free;
        end;
        If Pages.ActivePage.PageIndex = 1 then ChercherFormations;
end;

procedure TOF_PGCHANGEMENTPOSTE.MiseEnFormeGrille;
var GrilleEmploi,GrilleSalarie : THGrid;
    HMTrad: THSystemMenu;
    i : Integer;
    Libelle : String;
begin
        GrilleEmploi := THGrid(GetControl('GEMPLOI'));
        If GrilleEmploi <> Nil then
        begin
                GrilleEmploi.ColFormats[0] := 'CB=PGRHCOMPETENCES';
                GrilleEmploi.ColFormats[1] := '# ##0.00';
                GrilleEmploi.ColAligns[1] := taRightJustify;
                GrilleEmploi.ColFormats[2] := '# ##0.00';
                GrilleEmploi.ColAligns[2] := taRightJustify;
                GrilleEmploi.ColDrawingModes[3]:= 'IMAGE';
                GrilleEmploi.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
                For i := 1 to 5 do
                begin
                        Libelle := RechDom('GCZONELIBRE','RH'+IntToStr(i),False);
                        If Libelle <> 'Table libre '+IntToStr(i) then
                        begin
                                GrilleEmploi.ColCount := GrilleEmploi.ColCount + 1;
                                GrilleEmploi.ColFormats[GrilleEmploi.ColCount-1] := 'CB=PGTABLELIBRERH'+IntToStr(i);
                                GrilleEmploi.Titres.Add(Libelle);
                                GrilleEmploi.UpdateTitres;
                        end;
                end;
        end;
        GrilleSalarie := THGrid(GetControl('GSALARIE'));
        If GrilleSalarie <> Nil then
        begin
                GrilleSalarie.ColFormats[0] := 'CB=PGRHCOMPETENCES';
                GrilleSalarie.ColFormats[1] := '# ##0.00';
                GrilleSalarie.ColAligns[1] := taRightJustify;
                For i := 1 to 5 do
                begin
                        Libelle := RechDom('GCZONELIBRE','RH'+IntToStr(i),False);
                        If Libelle <> 'Table libre '+IntToStr(i) then
                        begin
                                GrilleSalarie.ColCount := GrilleSalarie.ColCount + 1;
                                GrilleSalarie.ColFormats[GrilleSalarie.ColCount-1] := 'CB=PGTABLELIBRERH'+IntToStr(i);
                                GrilleSalarie.Titres.Add(Libelle);
                                GrilleSalarie.UpdateTitres;
                        end;
                end;
        end;

        GrilleCompSal.ColFormats[2] := '# ##0.00';
        GrilleCompSal.ColAligns[2] := taRightJustify;
        GrilleCompSal.ColFormats[3] := '# ##0.00';
        GrilleCompSal.ColAligns[3] := taRightJustify;
        GrilleRecherche.ColFormats[2] := '# ##0.00';
        GrilleRecherche.ColAligns[2] := taRightJustify;
        GrilleRecherche.ColFormats[4] := '# ##0.00';
        GrilleRecherche.ColAligns[4] := taRightJustify;
        GrilleRecherche.ColEditables[0] := False;
        GrilleRecherche.ColEditables[1] := False;
        GrilleRecherche.ColEditables[3] := False;
        GrilleRecherche.ColEditables[4] := False;
        GrilleForm.ColFormats[2] := '# ##0.00';
        GrilleForm.ColAligns[2] := taRightJustify;
        GrilleForm.ColDrawingModes[3]:= 'IMAGE';  //PT2
        GrilleForm.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
        GrilleForm.ColFormats[4] := '# ##0.00';
        GrilleForm.ColAligns[4] := taRightJustify;
        GrilleRecherche.ColDrawingModes[3]:= 'IMAGE';
        GrilleRecherche.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
        GrilleSalarie.RowCount := 2;
        GrilleEmploi.RowCount := 2;
        GrilleRecherche.RowCount := 2;
        GrilleForm.RowCount := 2;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GEMPLOI'))) ;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GSALARIE'))) ;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPSAL'))) ;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPETENCES'))) ;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GRECHERCHE'))) ;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GFORMATIONS'))) ;
end;

procedure TOF_PGCHANGEMENTPOSTE.RetaillerGrille(Sender : TObject);
var HMTrad: THSystemMenu;
begin
  HMTrad.ResizeGridColumns(THGrid(GetControl('GEMPLOI'))) ;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GSALARIE'))) ;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPSAL'))) ;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPETENCES'))) ;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GRECHERCHE'))) ;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GFORMATIONS'))) ;
end;


procedure TOF_PGCHANGEMENTPOSTE.RemplirLesGrilles;
var Q : TQuery;
    TobCompetences,TC : Tob;
    HMTrad: THSystemMenu;
    Grille : THGrid;
    i : Integer;
begin
        GrilleCompSal := THGrid(GetControl('GCOMPSAL'));
        For i := 1 to GrilleCompSal.RowCount -1 do
        begin
                GrilleCompSal.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT PCH_COMPETENCE,PCO_LIBELLE,PCH_DEGREMAITRISE,'+
        '(SELECT POS_DEGREMAITRISE FROM STAGEOBJECTIF WHERE POS_CODESTAGE=PSA_LIBELLEEMPLOI AND POS_COMPETENCE=PCH_COMPETENCE) DEGREMAITRISEEMP '+
        'FROM RHCOMPETRESSOURCE '+
        'LEFT JOIN RHCOMPETENCES ON PCO_COMPETENCE=PCH_COMPETENCE '+
        'LEFT JOIN SALARIES ON PSA_SALARIE=PCH_SALARIE '+
        'WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="'+LeSalarie+'"',True);
        TobCompetences := Tob.Create('LesCompetences',Nil,-1);
        TobCompetences.LoadDetailDB('LesCompetences','','',Q,False);
        Ferme(Q);
        TobCompetences.PutGridDetail(GrilleCompSal,False,True,'',False);
        HMTrad.ResizeGridColumns(GrilleCompSal) ;
        If TobCompetences.Detail.count > 0 then GrilleCompSal.RowCount := TobCompetences.Detail.count + 1
        else GrilleCompSal.RowCount := 2;
        TobCompetences.Free;
        HMTrad.ResizeGridColumns(GrilleCompSal) ;
        For i := 1 to GrilleComp.RowCount -1 do
        begin
                GrilleComp.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT PCO_COMPETENCE,PCO_LIBELLE FROM RHCOMPETENCES',True);
        TobCompetences := Tob.Create('LesCompetences',Nil,-1);
        TobCompetences.LoadDetailDB('LesCompetences','','',Q,False);
        Ferme(Q);
        For i := 1 to GrilleCompSal.RowCount - 1 do
        begin
                TC := TobCompetences.FindFirst(['PCO_COMPETENCE'],[GrilleCompSal.CellValues[0,i]],False);
                If TC <> Nil then TC.Free;
        end;
        TobCompetences.PutGridDetail(GrilleComp,False,True,'',False);
        If TobCompetences.Detail.count > 0 then GrilleComp.RowCount := TobCompetences.Detail.count + 1
        else GrilleComp.RowCount := 2;
        TobCompetences.Free;
        HMTrad.ResizeGridColumns(GrilleComp) ;
end;

procedure TOF_PGCHANGEMENTPOSTE.GrilleDblCilck (Sender : TObject);
var i,NouvelleLigne : Integer;
begin
        If Sender = Nil then Exit;
        If THGrid(Sender) = GrilleComp then
        begin
                If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
                GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
                GrilleRecherche.FixedRows := 1;
                NouvelleLigne := GrilleRecherche.RowCount - 1;
                GrilleRecherche.CellValues[0,NouvelleLigne] := GrilleComp.CellValues[0,GrilleComp.Row];
                GrilleRecherche.CellValues[1,NouvelleLigne] := GrilleComp.CellValues[1,GrilleComp.Row];
                GrilleRecherche.CellValues[2,NouvelleLigne] := '0';
        end;
        If THGrid(Sender) = GrilleCompSal then
        begin
                If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
                GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
                GrilleRecherche.FixedRows := 1;
                NouvelleLigne := GrilleRecherche.RowCount - 1;
                GrilleRecherche.CellValues[0,NouvelleLigne] := GrilleCompSal.CellValues[0,GrilleCompSal.Row];
                GrilleRecherche.CellValues[1,NouvelleLigne] := GrilleCompSal.CellValues[1,GrilleCompSal.Row];
                GrilleRecherche.CellValues[2,NouvelleLigne] := '0';
        end;
end;

{procedure TOF_PGCHANGEMENTPOSTE.DeposeCompet;
var i,NouvelleLigne : Integer;
begin
        For i := 1 To GrilleComp.RowCount - 1 do
        begin
                If GrilleComp.IsSelected(i) then
                begin
                        If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
                        GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
                        GrilleRecherche.FixedRows := 1;
                        NouvelleLigne := GrilleRecherche.RowCount - 1;
                        GrilleRecherche.CellValues[0,NouvelleLigne] := GrilleComp.CellValues[0,i];
                        GrilleRecherche.CellValues[1,NouvelleLigne] := GrilleComp.CellValues[1,i];
                        GrilleRecherche.CellValues[2,NouvelleLigne] := '0';
                end;
        end;
        GrilleComp.ClearSelected;
end;

procedure TOF_PGCHANGEMENTPOSTE.DeposeCompetSal;
var i,NouvelleLigne : Integer;
begin
        For i := 1 To GrilleCompSal.RowCount - 1 do
        begin
                If GrilleCompSal.IsSelected(i) then
                begin
                        If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
                        GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
                        GrilleRecherche.FixedRows := 1;
                        NouvelleLigne := GrilleRecherche.RowCount - 1;
                        GrilleRecherche.CellValues[0,NouvelleLigne] := GrilleCompSal.CellValues[0,i];
                        GrilleRecherche.CellValues[1,NouvelleLigne] := GrilleCompSal.CellValues[1,i];
                        GrilleRecherche.CellValues[2,NouvelleLigne] := GrilleCompSal.CellValues[2,i];
                end;
        end;
        GrilleCompSal.ClearSelected;
end;}

Procedure TOF_PGCHANGEMENTPOSTE.GRechercheOnDblClick(Sender : TObject);
var i : Integer;
begin
        i := GrilleRecherche.Row;
        If (i = 1) and (GrilleRecherche.RowCount = 2) then GrilleRecherche.Rows[1].Clear
        Else GrilleRecherche.DeleteRow(i);
end;

procedure TOF_PGCHANGEMENTPOSTE.ChercherFormations;
var Q : TQuery;
    TobFormations : Tob;
    i,Lignes : Integer;
    StWhere : String;
    HMTrad: THSystemMenu;
begin
        GrilleForm.ClearSelected;
        For i := 1 to GrilleForm.RowCount -1 do
        begin
                GrilleForm.Rows[i].Clear;
        end;
        Lignes := grilleRecherche.RowCount - 1;
        StWhere := '';
        For i := 1 to Lignes do
        begin
                If grilleRecherche.CellValues[0,i] <> '' then
                begin
                        If StWhere = '' then StWhere := 'POS_COMPETENCE="'+grilleRecherche.CellValues[0,i]+'"'
                        else StWhere := StWhere + ' OR POS_COMPETENCE="'+grilleRecherche.CellValues[0,i]+'"';
                end;
        end;
        If StWhere = '' then Exit;
        Q := OpenSQL('SELECT PST_CODESTAGE,PST_LIBELLE,COUNT (POS_COMPETENCE) NBCOMPETENCES FROM STAGE LEFT JOIN STAGEOBJECTIF ON PST_CODESTAGE=POS_CODESTAGE '+
        'WHERE PST_MILLESIME="0000" AND ('+StWhere+') '+
        'group by PST_CODESTAGE,PST_LIBELLE',True);
        TobFormations := Tob.Create('LesCompetences',Nil,-1);
        TobFormations.LoadDetailDB('LesCompetences','','',Q,False);
        Ferme(Q);
        TobFormations.PutGridDetail(GrilleForm,False,True,'',False);
        If TobFormations.Detail.count > 0 then GrilleForm.RowCount := TobFormations.Detail.count + 1
        else GrilleForm.RowCount := 2;
        TobFormations.Free;
        HMTrad.ResizeGridColumns(GrilleForm) ;
        AfficherSelection(GrilleForm);
end;

procedure TOF_PGCHANGEMENTPOSTE.AfficherSelection(Sender : TObject);
var i,c,Ligne : Integer;
    TobComp : Tob;
    Q : TQuery;
    HMTrad: THSystemMenu;
begin
        For i := 1 to GrilleRecherche.RowCount -1 do
        begin
                GrilleRecherche.CellValues[3,i] := 'NON';
                GrilleRecherche.CellValues[4,i] := '0';
        end;
        Ligne := GrilleForm.Row;
        GrilleForm.ClearSelected;
        If GrilleForm.cellValues[0,Ligne] <> '' then
        begin
                Q := OpenSQL('SELECT POS_COMPETENCE,POS_DEGREMAITRISE FROM STAGEOBJECTIF WHERE POS_CODESTAGE="'+GrilleForm.cellValues[0,Ligne]+'"',True);
                TobComp := Tob.Create('LesComp',Nil,-1);
                TobComp.LoadDetailDB('LesComp','','',Q,False);
                Ferme(Q);
                For c := 0 to TobComp.Detail.Count - 1 do
                begin
                        For i := 1 to GrilleRecherche.RowCount -1 do
                        begin
                                If GrilleRecherche.CellValues[0,i] = TobComp.Detail[c].GetValue('POS_COMPETENCE') then
                                begin
                                        GrilleRecherche.CellValues[3,i] := 'OUI';
                                        GrilleRecherche.CellValues[4,i] := FloatToStr(TobComp.Detail[c].GetValue('POS_DEGREMAITRISE'));
                                end;
                        end;
                end;
                TobComp.Free;
        end
        else
        begin
                For i := 1 to GrilleRecherche.RowCount -1 do
                begin
                        GrilleRecherche.CellValues[3,i] := '';
                        GrilleRecherche.CellValues[4,i] := '0';
                end;
        end;
        HMTrad.ResizeGridColumns(GrilleRecherche) ;
end;

//DEBUT PT2
procedure TOF_PGCHANGEMENTPOSTE.AfficherSelectionComp(Sender : TObject);
var i,c,Ligne : Integer;
    TobComp : Tob;
    Q : TQuery;
    HMTrad: THSystemMenu;
begin
        For i := 1 to GrilleForm.RowCount -1 do
        begin
                GrilleForm.CellValues[3,i] := 'NON';
                GrilleForm.CellValues[4,i] := '0';
        end;
        Ligne := GrilleRecherche.Row;
        GrilleRecherche.ClearSelected;
        If GrilleRecherche.cellValues[0,Ligne] <> '' then
        begin
                Q := OpenSQL('SELECT POS_CODESTAGE,POS_DEGREMAITRISE FROM STAGEOBJECTIF WHERE POS_COMPETENCE="'+GrilleRecherche.cellValues[0,Ligne]+'"',True);
                TobComp := Tob.Create('Lesformations',Nil,-1);
                TobComp.LoadDetailDB('Lesformations','','',Q,False);
                Ferme(Q);
                For c := 0 to TobComp.Detail.Count - 1 do
                begin
                        For i := 1 to GrilleForm.RowCount -1 do
                        begin
                                If GrilleForm.CellValues[0,i] = TobComp.Detail[c].GetValue('POS_CODESTAGE') then
                                begin
                                        GrilleForm.CellValues[3,i] := 'OUI';
                                        GrilleForm.CellValues[4,i] := FloatToStr(TobComp.Detail[c].GetValue('POS_DEGREMAITRISE'));
                                end;
                        end;
                end;
                TobComp.Free;
        end
        else
        begin
                For i := 1 to GrilleForm.RowCount -1 do
                begin
                        GrilleForm.CellValues[3,i] := '';
                        GrilleForm.CellValues[4,i] := '0';
                end;
        end;
        HMTrad.ResizeGridColumns(GrilleForm) ;
end;

//FIN PT2

procedure TOF_PGCHANGEMENTPOSTE.ChangeOnglet (Sender : TObject);
var GEmploi : THGrid;
    i,NouvelleLigne : Integer;
    HMTrad: THSystemMenu;
begin
        For i := 1 to GrilleRecherche.RowCount -1 do
        begin
                GrilleRecherche.Rows[i].Clear;
        end;
        GrilleRecherche.RowCount := 2;
        GEmploi := THGrid(GetControl('GEMPLOI'));
        For i := 1 to GEmploi.RowCount - 1 do
        begin
                If GEmploi.Cellvalues[3,i] = 'NON' then
                begin
                        If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
                        GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
                        GrilleRecherche.FixedRows := 1;
                        NouvelleLigne := GrilleRecherche.RowCount - 1;
                        GrilleRecherche.CellValues[0,NouvelleLigne] := GEmploi.CellValues[0,i];
                        GrilleRecherche.CellValues[1,NouvelleLigne] := RechDom('PGRHCOMPETENCES',GEmploi.CellValues[0,i],False);
                        GrilleRecherche.CellValues[2,NouvelleLigne] := GEmploi.CellValues[1,i];
                end;
        end;
        If TTabSheet(Sender).Name = 'PCOMPETENCES' then ChercherFormations;
        HMTrad.ResizeGridColumns(GrilleRecherche) ;
end;

{procedure TOF_PGCHANGEMENTPOSTE.GrilleMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
        If not (ssCtrl in Shift) then
        begin
                If Sender = GrilleComp then GrilleComp.BeginDrag(TRUE, 10);
                If Sender = GrilleCompSal then GrilleCompSal.BeginDrag(TRUE, 10);
        end;
end;

procedure TOF_PGCHANGEMENTPOSTE.DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);
begin
        If Source = GrilleComp then DeposeCompet;
        If Source = GrilleCompSal then DeposeCompetSal;
end;}

{procedure TOF_PGCHANGEMENTPOSTE.TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := FALSE;
  if (Sender is THgrid) then Accept := TRUE;
end;}

Procedure TOF_PGCHANGEMENTPOSTE.ZoomChoixCompet(Sender : TObject);
var ListeCompet,Competence : String;
    NouvelleLigne : Integer;
begin
        ListeCompet := AGLLanceFiche('PAY','PGMUL_RHCOMPET','','','SIMULATIONPOSTE');
        While ListeCompet <> '' do
        begin
                Competence := ReadTokenPipe(ListeCompet,';');
                If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
                GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
                GrilleRecherche.FixedRows := 1;
                NouvelleLigne := GrilleRecherche.RowCount - 1;
                GrilleRecherche.CellValues[0,NouvelleLigne] := Competence;
                GrilleRecherche.CellValues[1,NouvelleLigne] := Competence;
                GrilleRecherche.CellValues[2,NouvelleLigne] := '0';
        end;
end;

procedure TOF_PGCHANGEMENTPOSTE.BImprimerClick(Sender : TOBject);
begin
        ImpressionSimulation(False);
end;


procedure TOF_PGCHANGEMENTPOSTE.ImpressionSimulation(Exporter : Boolean);
var TobEtat,T,TobCompet,TobCompetForm,TF,TobFormations,TC : Tob;
    Q : TQuery;
    Pages : TPageControl;
    i,x : Integer;
    NiveauCompet : Double;
    Formation,Competence,SommeCompet,Salarie : String;
    Prenom,NomSal,LibEmploi,LibEmploiFut,Etablissement,Travailn1,Travailn2,Travailn3,Travailn4 : String;
    DateEntree : TDateTime;
begin
        SetControlChecked('EXPORT',Exporter);
        Pages := TPageControl(Getcontrol('Pages'));
        Q := OpenSQL('SELECT POS_NATOBJSTAGE,POS_CODESTAGE,POS_ORDRE,POS_COMPETENCE,'+
        'POS_DEGREMAITRISE,POS_TABLELIBRECR1,POS_TABLELIBRECR2,POS_TABLELIBRECR3,'+
        'PST_DUREESTAGE,PST_COUTBUDGETE,PST_NATUREFORM,PST_CENTREFORMGU,PST_FORMATION1,PST_FORMATION2,'+
        'PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8 '+
        'FROM STAGEOBJECTIF '+
        'LEFT JOIN STAGE ON POS_CODESTAGE=PST_CODESTAGE '+
        'WHERE POS_NATOBJSTAGE="FOR" AND PST_MILLESIME="0000"',True);
        TobCompetForm := Tob.Create('lescompetences',Nil,-1);
        TobCompetForm.LoadDetailDB('lescompetences','','',Q,False);
        Ferme(Q);
        TobFormations := Tob.Create('lesformations',Nil,-1);
        For i := 1 to GrilleForm.RowCount -1 do
        begin
                Formation := GrilleForm.CellValues[0,i];
                SommeCompet := '';
                For x := 1 to GrilleRecherche.RowCount -1 do
                begin
                        Competence := GrilleRecherche.CellValues[0,x];
                        TC := TobCompetForm.FindFirst(['POS_CODESTAGE','POS_COMPETENCE'],[Formation,Competence],False);
                        If TC <> Nil then
                        begin
                                TF := TOb.create('FilleForm',TobFormations,-1);
                                TF.AddChampSupValeur('TYPECRITERE','SELECTION');
                                TF.AddChampSupValeur('LESTAGE',Formation);
                                TF.AddChampSupValeur('COMPETENCE',Competence);
                                TF.AddChampSupValeur('PST_DUREESTAGE',TC.GetValue('PST_DUREESTAGE'));
                                TF.AddChampSupValeur('PST_COUTBUDGETE',TC.GetValue('PST_COUTBUDGETE'));
                                TF.AddChampSupValeur('PST_NATUREFORM',TC.GetValue('PST_NATUREFORM'));
                                TF.AddChampSupValeur('PST_CENTREFORMGU',TC.GetValue('PST_CENTREFORMGU'));
                                TF.AddChampSupValeur('PST_FORMATION1',TC.GetValue('PST_FORMATION1'));
                                TF.AddChampSupValeur('PST_FORMATION2',TC.GetValue('PST_FORMATION2'));
                                TF.AddChampSupValeur('PST_FORMATION3',TC.GetValue('PST_FORMATION3'));
                                TF.AddChampSupValeur('PST_FORMATION4',TC.GetValue('PST_FORMATION4'));
                                TF.AddChampSupValeur('PST_FORMATION5',TC.GetValue('PST_FORMATION5'));
                                TF.AddChampSupValeur('PST_FORMATION6',TC.GetValue('PST_FORMATION6'));
                                TF.AddChampSupValeur('PST_FORMATION7',TC.GetValue('PST_FORMATION7'));
                                TF.AddChampSupValeur('PST_FORMATION8',TC.GetValue('PST_FORMATION8'));
                                TF.AddChampSupValeur('POS_DEGREMAITRISE',TC.GetValue('POS_DEGREMAITRISE'));
                                TF.AddChampSupValeur('POS_TABLELIBRECR1',TC.GetValue('POS_TABLELIBRECR1'));
                                TF.AddChampSupValeur('POS_TABLELIBRECR2',TC.GetValue('POS_TABLELIBRECR2'));
                                TF.AddChampSupValeur('POS_TABLELIBRECR3',TC.GetValue('POS_TABLELIBRECR3'));
                        end;
                end;
        end;
        TobCompetForm.Free;
        Salarie := GetControltext('SALARIE');
        Q := OpenSQL('SELECT * FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);
        If Not Q.Eof then
        begin
                Prenom := Q.FindField('PSA_PRENOM').AsString;
                NomSal := Q.FindField('PSA_LIBELLE').AsString;
                LibEmploi := Q.FindField('PSA_LIBELLEEMPLOI').AsString;
                If LibEmploi <> '' then LibEmploi := RechDom('PGLIBEMPLOI',LibEmploi,False);
                Etablissement := Q.FindField('PSA_ETABLISSEMENT').AsString;
                If Etablissement <> '' then Etablissement := RechDom('TTETABLISSEMENT',Etablissement,False);
                Travailn1 := Q.FindField('PSA_TRAVAILN1').AsString;
                If Travailn1 <> '' then Travailn1 := RechDom('PGTRAVAILN1',Travailn1,False);
                Travailn2 := Q.FindField('PSA_TRAVAILN2').AsString;
                If Travailn2 <> '' then Travailn2 := RechDom('PGTRAVAILN2',Travailn2,False);
                Travailn3 := Q.FindField('PSA_TRAVAILN3').AsString;
                If Travailn3 <> '' then Travailn3 := RechDom('PGTRAVAILN3',Travailn3,False);
                Travailn4 := Q.FindField('PSA_TRAVAILN4').AsString;
                If Travailn4 <> '' then Travailn4 := RechDom('PGTRAVAILN4',Travailn4,False);
                DateEntree := Q.FindField('PSA_DATEENTREE').AsDateTime;
        end;
        Ferme(Q);
        LibEmploiFut := GetControlText('LIBEMPLOIFUT');
        TobEtat := Tob.Create('Edition',Nil,-1);
        TobEtat.AddChampSup('PSA_SALARIE',False);
        TobEtat.AddChampSup('TYPECRITERE',False);
        TobEtat.AddChampSup('LESTAGE',False);
        TobEtat.AddChampSup('COMPETENCE',False);
        TobEtat.AddChampSup('ETABLISSEMENT',False);
        TobEtat.AddChampSup('TRAVAILN1',False);
        TobEtat.AddChampSup('TRAVAILN2',False);
        TobEtat.AddChampSup('TRAVAILN3',False);
        TobEtat.AddChampSup('TRAVAILN4',False);
        TobEtat.AddChampSup('LIBEMPLOI',False);
        TobEtat.AddChampSup('LIBEMPLOIFUT',False);
        TobEtat.AddChampSup('DATEENTREE',False);
        TobEtat.AddChampSup('LIBELLE',False);
        TobEtat.AddChampSup('PRENOM',False);
        TobEtat.AddChampSup('LIBSTAGE',False);
        TobEtat.AddChampSup('LIBCOMPETENCE',False);
        TobEtat.AddChampSup('PST_DUREESTAGE',False);
        TobEtat.AddChampSup('PST_COUTBUDGETE',False);
        TobEtat.AddChampSup('PST_NATUREFORM',False);
        TobEtat.AddChampSup('PST_CENTREFORMGU',False);
        TobEtat.AddChampSup('PST_FORMATION1',False);
        TobEtat.AddChampSup('PST_FORMATION2',False);
        TobEtat.AddChampSup('PST_FORMATION3',False);
        TobEtat.AddChampSup('PST_FORMATION4',False);
        TobEtat.AddChampSup('PST_FORMATION5',False);
        TobEtat.AddChampSup('PST_FORMATION6',False);
        TobEtat.AddChampSup('PST_FORMATION7',False);
        TobEtat.AddChampSup('PST_FORMATION8',False);
        TobEtat.AddChampSup('POS_DEGREMAITRISE',False);
        TobEtat.AddChampSup('POS_TABLELIBRECR1',False);
        TobEtat.AddChampSup('POS_TABLELIBRECR2',False);
        TobEtat.AddChampSup('POS_TABLELIBRECR3',False);
        TobEtat.AddChampSup('LIBRERH1',False);
        TobEtat.AddChampSup('LIBRERH2',False);
        TobEtat.AddChampSup('LIBRERH3',False);
        TobEtat.AddChampSup('LIBRERH4',False);
        TobEtat.AddChampSup('LIBRERH5',False);
        Q := OpenSQL('SELECT POS_NATOBJSTAGE,POS_CODESTAGE,POS_ORDRE,POS_COMPETENCE,'+
        'POS_DEGREMAITRISE,POS_TABLELIBRECR1,POS_TABLELIBRECR2,POS_TABLELIBRECR3,'+
        'PCO_TABLELIBRERH1,PCO_TABLELIBRERH2,PCO_TABLELIBRERH3,PCO_TABLELIBRERH4,PCO_TABLELIBRERH5 '+
        'FROM STAGEOBJECTIF '+
        'LEFT JOIN RHCOMPETENCES ON POS_COMPETENCE=PCO_COMPETENCE '+
        'WHERE POS_NATOBJSTAGE="EMP" AND POS_CODESTAGE="'+GetControlText('EMPLOIFUT')+'"',True);
        TobCompetForm := Tob.Create('lescompetences',Nil,-1);
        TobCompetForm.LoadDetailDB('lescompetences','','',Q,False);
        Ferme(Q);
        For i := 1 to GrilleRecherche.RowCount -1 do
        begin
                Competence := GrilleRecherche.CellValues[0,i];
                If IsNumeric(GrilleRecherche.CellValues[2,x]) then NiveauCompet := StrToFloat(GrilleRecherche.CellValues[2,x])
                else NiveauCompet := 0;
                T := Tob.Create('FilleEtat',TobEtat,-1);
                T.AddChampSupValeur('PSA_SALARIE',Salarie);
                T.AddChampSupValeur('TYPECRITERE','1');
                T.AddChampSupValeur('LESTAGE','');
                T.AddChampSupValeur('COMPETENCE',Competence);
                T.AddChampSupValeur('ETABLISSEMENT',Etablissement);
                T.AddChampSupValeur('TRAVAILN1',Travailn1);
                T.AddChampSupValeur('TRAVAILN2',Travailn2);
                T.AddChampSupValeur('TRAVAILN3',Travailn3);
                T.AddChampSupValeur('TRAVAILN4',Travailn4);
                T.AddChampSupValeur('LIBEMPLOI',LibEmploi);
                T.AddChampSupValeur('LIBEMPLOIFUT',LibEmploiFut);
                T.AddChampSupValeur('DATEENTREE',DateEntree);
                T.AddChampSupValeur('LIBELLE',NomSal);
                T.AddChampSupValeur('PRENOM',Prenom);
                T.AddChampSupValeur('LIBSTAGE','');
                T.AddChampSupValeur('LIBCOMPETENCE',RechDom('PGRHCOMPETENCES',Competence,False));
                T.AddChampSupValeur('PST_DUREESTAGE',0);
                T.AddChampSupValeur('PST_COUTBUDGETE',0);
                T.AddChampSupValeur('PST_NATUREFORM','');
                T.AddChampSupValeur('PST_CENTREFORMGU','');
                T.AddChampSupValeur('PST_FORMATION1','');
                T.AddChampSupValeur('PST_FORMATION2','');
                T.AddChampSupValeur('PST_FORMATION3','');
                T.AddChampSupValeur('PST_FORMATION4','');
                T.AddChampSupValeur('PST_FORMATION5','');
                T.AddChampSupValeur('PST_FORMATION6','');
                T.AddChampSupValeur('PST_FORMATION7','');
                T.AddChampSupValeur('PST_FORMATION8','');
                T.AddChampSupValeur('POS_DEGREMAITRISE',NiveauCompet);
                TC := TobCompetForm.FindFirst(['POS_COMPETENCE'],[Competence],False);
                If TC <> Nil then
                begin
                        T.AddChampSupValeur('POS_TABLELIBRECR1',TC.GetValue('POS_TABLELIBRECR1'));
                        T.AddChampSupValeur('POS_TABLELIBRECR2',TC.GetValue('POS_TABLELIBRECR2'));
                        T.AddChampSupValeur('POS_TABLELIBRECR3',TC.GetValue('POS_TABLELIBRECR3'));
                        T.AddChampSupValeur('LIBRERH1',RechDom('PGTABLELIBRERH1',TC.GetValue('PCO_TABLELIBRERH1'),False));
                        T.AddChampSupValeur('LIBRERH2',RechDom('PGTABLELIBRERH2',TC.GetValue('PCO_TABLELIBRERH2'),False));
                        T.AddChampSupValeur('LIBRERH3',RechDom('PGTABLELIBRERH3',TC.GetValue('PCO_TABLELIBRERH3'),False));
                        T.AddChampSupValeur('LIBRERH4',RechDom('PGTABLELIBRERH4',TC.GetValue('PCO_TABLELIBRERH4'),False));
                        T.AddChampSupValeur('LIBRERH5',RechDom('PGTABLELIBRERH5',TC.GetValue('PCO_TABLELIBRERH5'),False));
                end
                else
                begin
                        T.AddChampSupValeur('POS_TABLELIBRECR1','');
                        T.AddChampSupValeur('POS_TABLELIBRECR2','');
                        T.AddChampSupValeur('POS_TABLELIBRECR3','');
                        T.AddChampSupValeur('LIBRERH1','');
                        T.AddChampSupValeur('LIBRERH2','');
                        T.AddChampSupValeur('LIBRERH3','');
                        T.AddChampSupValeur('LIBRERH4','');
                        T.AddChampSupValeur('LIBRERH5','');
                end;
        end;
        TobCompetForm.Free;
        For i := 0 to TobFormations.Detail.Count - 1 do
        begin
                T := Tob.Create('FilleEtat',TobEtat,-1);
                Formation := TobFormations.detail[i].GetValue('LESTAGE');
                Competence := TobFormations.detail[i].GetValue('COMPETENCE');
                T.AddChampSupValeur('PSA_SALARIE',Salarie);
                T.AddChampSupValeur('TYPECRITERE','2');
                T.AddChampSupValeur('LESTAGE',Formation);
                T.AddChampSupValeur('COMPETENCE',Competence);
                T.AddChampSupValeur('ETABLISSEMENT',Etablissement);
                T.AddChampSupValeur('TRAVAILN1',Travailn1);
                T.AddChampSupValeur('TRAVAILN2',Travailn2);
                T.AddChampSupValeur('TRAVAILN3',Travailn3);
                T.AddChampSupValeur('TRAVAILN4',Travailn4);
                T.AddChampSupValeur('LIBEMPLOI',LibEmploi);
                T.AddChampSupValeur('LIBEMPLOIFUT',LibEmploiFut);
                T.AddChampSupValeur('DATEENTREE',DateEntree);
                T.AddChampSupValeur('LIBELLE',NomSal);
                T.AddChampSupValeur('PRENOM',Prenom);
                T.AddChampSupValeur('LIBSTAGE',RechDom('PGSTAGEFORMCOMPLET',Formation,False));
                T.AddChampSupValeur('LIBCOMPETENCE',RechDom('PGRHCOMPETENCES',Competence,False));
                T.AddChampSupValeur('PST_DUREESTAGE',TobFormations.detail[i].GetValue('PST_DUREESTAGE'));
                T.AddChampSupValeur('PST_COUTBUDGETE',TobFormations.detail[i].GetValue('PST_COUTBUDGETE'));
                T.AddChampSupValeur('PST_NATUREFORM',TobFormations.detail[i].GetValue('PST_NATUREFORM'));
                T.AddChampSupValeur('PST_CENTREFORMGU',TobFormations.detail[i].GetValue('PST_CENTREFORMGU'));
                T.AddChampSupValeur('PST_FORMATION1',TobFormations.detail[i].GetValue('PST_FORMATION1'));
                T.AddChampSupValeur('PST_FORMATION2',TobFormations.detail[i].GetValue('PST_FORMATION2'));
                T.AddChampSupValeur('PST_FORMATION3',TobFormations.detail[i].GetValue('PST_FORMATION3'));
                T.AddChampSupValeur('PST_FORMATION4',TobFormations.detail[i].GetValue('PST_FORMATION4'));
                T.AddChampSupValeur('PST_FORMATION5',TobFormations.detail[i].GetValue('PST_FORMATION5'));
                T.AddChampSupValeur('PST_FORMATION6',TobFormations.detail[i].GetValue('PST_FORMATION6'));
                T.AddChampSupValeur('PST_FORMATION7',TobFormations.detail[i].GetValue('PST_FORMATION7'));
                T.AddChampSupValeur('PST_FORMATION8',TobFormations.detail[i].GetValue('PST_FORMATION8'));
                T.AddChampSupValeur('POS_DEGREMAITRISE',TobFormations.detail[i].GetValue('POS_DEGREMAITRISE'));
                T.AddChampSupValeur('POS_TABLELIBRECR1',TobFormations.detail[i].GetValue('POS_TABLELIBRECR1'));
                T.AddChampSupValeur('POS_TABLELIBRECR2',TobFormations.detail[i].GetValue('POS_TABLELIBRECR2'));
                T.AddChampSupValeur('POS_TABLELIBRECR3',TobFormations.detail[i].GetValue('POS_TABLELIBRECR3'));
                T.AddChampSupValeur('LIBRERH1','');
                T.AddChampSupValeur('LIBRERH2','');
                T.AddChampSupValeur('LIBRERH3','');
                T.AddChampSupValeur('LIBRERH4','');
                T.AddChampSupValeur('LIBRERH5','');
        end;
        TobFormations.Free;
        TobEtat.Detail.Sort('PSA_SALARIE;TYPECRITERE;LESTAGE');
        If GetCheckBoxState('EXPORT') = CbChecked then LanceEtatTOB('E','PCO','PL1',TobEtat,True,true,False,Pages,'','',False)
        else LanceEtatTOB('E','PCO','PL1',TobEtat,True,False,False,Pages,'','',False);
        TobEtat.Free;
end;

procedure TOF_PGCHANGEMENTPOSTE.BExportClick(Sender : Tobject);
begin
        ImpressionSimulation(True);
end;

Initialization
  registerclasses ( [ TOF_PGCHANGEMENTPOSTE ] ) ;
end.

