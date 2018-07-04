{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPROPOSITIONSTAGE ()
Mots clefs ... : TOF;PGPROPOSITIONSTAGE
*****************************************************************}
Unit UTofPGPropositionStage ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,uTob,HSysMenu,EntPaie,ParamSoc,HTB97,Vierge,P5DEF ;

Type
  TOF_PGPROPOSITIONSTAGE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    LeSalarie : String;
    GrilleComp,GrilleCompSal,GrilleRecherche,GrilleForm : THGrid;
    procedure RemplirLesGrilles;
    procedure GCompetenceSalDblClick(Sender : TObject);
    procedure GCompetenceDblClick(Sender : TObject);
    Procedure GRechercheOnDblClick(Sender : TObject);
    procedure ChercherFormations(Sender : TObject);
    procedure AfficherSelection(Sender : TObject);
  end ;

Implementation


procedure TOF_PGPROPOSITIONSTAGE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGPROPOSITIONSTAGE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGPROPOSITIONSTAGE.OnArgument (S : String ) ;
var Grille : THGrid;
    BCherche : TToolBarButton97;
begin
  Inherited ;
  LeSalarie := ReadTokenPipe(S,';');
  RemplirLesGrilles;
  GrilleComp := THGrid(GetControl('GCOMPETENCES'));
  GrilleComp.OnDblClick := GCompetenceDblClick;
  GrilleCompSal := THGrid(GetControl('GCOMPSAL'));
  GrilleCompSal.OnDblClick := GCompetenceSalDblClick;
  GrilleRecherche := THGrid(GetControl('GRECHERCHE'));
  GrilleRecherche.OnDblClick := GRechercheOnDblClick;
  GrilleRecherche.RowCount := 2;
  BCherche := TToolBarButton97(GetControl('BCHERCHE'));
  BCherche.OnClick := ChercherFormations;
  Ecran.Caption := 'Proposition de formations pour le salarié '+ LeSalarie + ' ' + RechDom('PGSALARIE',LeSalarie,False);
  UpdateCaption(Ecran);
  GrilleForm := THGrid(GetControl('GFORMATIONS'));
  GrilleForm.OnClick := AfficherSelection;
  GrilleRecherche.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
end ;

procedure TOF_PGPROPOSITIONSTAGE.RemplirLesGrilles;
var Q : TQuery;
    TobCompetences : Tob;
    HMTrad: THSystemMenu;
    Grille : THGrid;
    i : Integer;
begin
        Grille := THGrid(GetControl('GCOMPSAL'));
        For i := 1 to Grille.RowCount -1 do
        begin
                Grille.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT PCH_COMPETENCE,PCO_LIBELLE,PCH_DEGREMAITRISE FROM RHCOMPETRESSOURCE '+
        'LEFT JOIN RHCOMPETENCES ON PCO_COMPETENCE=PCH_COMPETENCE '+
        'WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="'+LeSalarie+'"',True);
        TobCompetences := Tob.Create('LesCompetences',Nil,-1);
        TobCompetences.LoadDetailDB('LesCompetences','','',Q,False);
        Ferme(Q);
        TobCompetences.PutGridDetail(Grille,False,True,'',False);
        If TobCompetences.Detail.count > 0 then Grille.RowCount := TobCompetences.Detail.count + 1
        else Grille.RowCount := 2;
        TobCompetences.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPSAL'))) ;
        Grille := THGrid(GetControl('GCOMPETENCES'));
        For i := 1 to Grille.RowCount -1 do
        begin
                Grille.Rows[i].Clear;
        end;
        Q := OpenSQL('SELECT PCO_COMPETENCE,PCO_LIBELLE FROM RHCOMPETENCES',True);
        TobCompetences := Tob.Create('LesCompetences',Nil,-1);
        TobCompetences.LoadDetailDB('LesCompetences','','',Q,False);
        Ferme(Q);
        TobCompetences.PutGridDetail(Grille,False,True,'',False);
        If TobCompetences.Detail.count > 0 then Grille.RowCount := TobCompetences.Detail.count + 1
        else Grille.RowCount := 2;
        TobCompetences.Free;
        HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPETENCES'))) ;
end;

procedure TOF_PGPROPOSITIONSTAGE.GCompetenceDblClick(Sender : TObject);
var Ligne,NouvelleLigne : Integer;
begin
        Ligne := GrilleComp.Row;
        If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
        GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
        GrilleRecherche.FixedRows := 1;
        NouvelleLigne := GrilleRecherche.RowCount - 1;
        GrilleRecherche.CellValues[0,NouvelleLigne] := GrilleComp.CellValues[0,Ligne];
        GrilleRecherche.CellValues[1,NouvelleLigne] := GrilleComp.CellValues[1,Ligne];
        GrilleRecherche.CellValues[2,NouvelleLigne] := '0';
end;

procedure TOF_PGPROPOSITIONSTAGE.GCompetenceSalDblClick(Sender : TObject);
var Ligne,NouvelleLigne : Integer;
begin
        Ligne := GrilleCompSal.Row;
        If (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0,1] = '') then GrilleRecherche.RowCount := 1;
        GrilleRecherche.RowCount:=GrilleRecherche.RowCount+1;
        GrilleRecherche.FixedRows := 1;
        NouvelleLigne := GrilleRecherche.RowCount - 1;
        GrilleRecherche.CellValues[0,NouvelleLigne] := GrilleCompSal.CellValues[0,Ligne];
        GrilleRecherche.CellValues[1,NouvelleLigne] := GrilleCompSal.CellValues[1,Ligne];
        GrilleRecherche.CellValues[2,NouvelleLigne] := GrilleCompSal.CellValues[2,Ligne];
end;

Procedure TOF_PGPROPOSITIONSTAGE.GRechercheOnDblClick(Sender : TObject);
var i : Integer;
begin
        i := GrilleRecherche.Row;
        If (i = 1) and (GrilleRecherche.RowCount = 2) then GrilleRecherche.Rows[1].Clear
        Else GrilleRecherche.DeleteRow(i);
end;

procedure TOF_PGPROPOSITIONSTAGE.ChercherFormations(Sender : TObject);
var Q : TQuery;
    TobFormations : Tob;
    i,Lignes : Integer;
    StWhere : String;
begin
        For i := 1 to GrilleForm.RowCount -1 do
        begin
                GrilleForm.Rows[i].Clear;
        end;
        Lignes := grilleRecherche.RowCount - 1;
        StWhere := '';
        For i := 1 to Lignes do
        begin
                If GrilleForm.CellValues[0,i] <> '' then
                begin
                        If StWhere = '' then StWhere := ' WHERE POS_COMPETENCE="'+GrilleForm.CellValues[0,i]+'"'
                        else StWhere := StWhere + ' OR POS_COMPETENCE="'+GrilleForm.CellValues[0,i]+'"';
                end;
        end;
        Q := OpenSQL('SELECT PST_CODESTAGE,PST_LIBELLE,POS_DEGREMAITRISE FROM STAGE LEFT JOIN STAGEOBJECTIF ON PST_CODESTAGE=POS_CODESTAGE WHERE PST_MILLESIME="0000" AND '+
        'PST_CODESTAGE IN (SELECT POS_CODESTAGE FROM STAGEOBJECTIF'+StWhere+') '+
        'group by PST_CODESTAGE,PST_LIBELLE,POS_DEGREMAITRISE',True);
        TobFormations := Tob.Create('LesCompetences',Nil,-1);
        TobFormations.LoadDetailDB('LesCompetences','','',Q,False);
        Ferme(Q);
        TobFormations.PutGridDetail(GrilleForm,False,True,'',False);
        If TobFormations.Detail.count > 0 then GrilleForm.RowCount := TobFormations.Detail.count + 1
        else GrilleForm.RowCount := 2;
        TobFormations.Free;
//        HMTrad.ResizeGridColumns(THGrid(GetControl('GFORMATIONS'))) ;
end;

procedure TOF_PGPROPOSITIONSTAGE.AfficherSelection(Sender : TObject);
var i,c,Ligne : Integer;
    TobComp : Tob;
    Q : TQuery;
begin
        For i := 1 to GrilleRecherche.RowCount -1 do
        begin
                GrilleRecherche.CellValues[3,i] := 'NON';
                GrilleRecherche.CellValues[4,i] := '0';
        end;
        Ligne := GrilleForm.Row;
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
end;

Initialization
  registerclasses ( [ TOF_PGPROPOSITIONSTAGE ] ) ;
end.

