{***********UNITE*************************************************
Auteur  ...... : FL
Créé le ...... : 03/05/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULSTAGEEMPLOI
Mots clefs ... : TOF;MULSTAGEEMPLOI
*****************************************************************
PT1 | 15/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT2 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT3 | 17/04/2008 | V_803 | FL | Prise en compte du mode d'ouverture en lecture/écriture
PT4 | 17/04/2008 | V_803 | FL | Ajout de PGBundleCatalogue
}
Unit UTOFPGMulStageEmploi;

Interface

Uses Controls,Classes,sysutils,ComCtrls,HTB97,UTOF,
{$IFNDEF EAGLCLIENT}
      HDB,Mul,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     emul,eFiche,UtileAGL,
{$ENDIF}
     HMsgBox,HCtrls,HEnt1,Hqry,UTOB,LookUp, ParamSoc, StdCtrls,PGOutilsFormation,SaisieList;

Type
  TOF_MULSTAGEEMPLOI = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    Private
    LeStage,Ordre : String;
    TobDispo,TobAssoc : TOB;
    GrilleDepart,GrilleDest : THGrid;
    DragRow : Integer;

    Procedure BAjoutClick(Sender : TObject);
    Procedure BSuppClick(Sender : TObject);
    Procedure AfficheGrille;
    procedure Depose_Objet(Destination, Origine : TObject; X,Y : Integer);
    procedure AssocToDispo(i: integer);
    procedure DispoToAssoc(i: integer);
    procedure OnDragOver (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure OnMouseDown(Sender : TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  end ;

  TOF_LISTESTAGEEMPLOI = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ; //PT1
  end ;

Implementation

Uses GalOutil;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Effectue un annule et remplace des emplois associés au 
Suite ........ : stage
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.OnUpdate ;
var i : Integer;
    TobMaj,T : Tob;
begin
  Inherited ;
  Try
     BEGINTRANS;
     
     ExecuteSQL('DELETE FROM LIENEMPFORM WHERE PMF_CODESTAGE="'+LeStage+'" AND PMF_ORDRE="-1"');

     TobMaj := TOb.Create('£EmploisAssociés', Nil, -1);
     For i:= 0 to TobAssoc.Detail.Count - 1 Do
     Begin
          T := Tob.Create('LIENEMPFORM', TobMaj, -1);
          T.PutValue('PMF_CODESTAGE', LeStage);
          T.PutValue('PMF_ORDRE',     '-1');
          T.PutValue('PMF_LIBELLEEMPLOI', TobAssoc.Detail[i].GetValue('CC_CODE'));
     end;

     TobMaj.InsertDB(Nil, False);

     COMMITTRANS;
  Except
     ROLLBACK;
  End;
  FreeAndNil(TobMaj);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Chargement des données à l'ouverture de l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.OnLoad ;
var
    SQL : String;
    T   : TOB;
    Q   : TQuery;
begin
  Inherited ;
     If TobDispo <> Nil Then FreeAndNil(TobDispo);
     If TobAssoc <> Nil Then FreeAndNil(TobAssoc);

     // Récupération des codes emplois
     TobDispo := TOB.Create('LesDispos', Nil, -1);
     SQL := 'SELECT CC_CODE,CC_LIBELLE FROM ';
     // Gestion du mutidossiers
     If PGBundleInscFormation Then
     Begin
          // Recherche de la tablette correspondante
          Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="PGLIBEMPLOI" AND DS_TYPTABLE="TTE"', True);
          If Not Q.EOF Then SQL := SQL + Q.FindField('DS_NOMBASE').AsString + '.DBO.';
          Ferme(Q);
     End;
     SQL := SQL + 'CHOIXCOD WHERE CC_TYPE="PLE"';
     TobDispo.LoadDetailFromSQL(SQL);

     // Mise à jour de la TOB d'association avec les libellés emplois déjà affectés au stage
     TobAssoc := TOB.Create('LesAssociés', Nil, -1);
     Q := OpenSQL('SELECT PMF_LIBELLEEMPLOI FROM LIENEMPFORM WHERE PMF_CODESTAGE="'+LeStage+'" AND PMF_ORDRE="-1"', True);
     While Not Q.EOF Do
     Begin
          T := TobDispo.FindFirst(['CC_CODE'], [Q.FindField('PMF_LIBELLEEMPLOI').AsString], False);
          If T <> Nil Then T.ChangeParent(TobAssoc, -1);
          Q.Next;
     End;
     Ferme(Q);

     // Affiche les données dans les grilles
     AfficheGrille;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Ouverture de l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.OnArgument (S : String ) ;
var BtPlus,BtMoins : TToolBarButton97;
    Action : String;
begin
  Inherited ;
     LeStage := ReadTokenSt(S);
     Ordre   := ReadTokenSt(S);
     Action  := ReadTokenSt(S);

     // Récupération des grilles dispo et assoc
     GrilleDepart := THGrid(GetControl('GDISPO'));
     GrilleDest   := THGrid(GetControl('GASSOC'));

     // Boutons d'ajout et de suppression des emplois associés
     BtPlus := TToolBarButton97(GetControl('BAJOUT'));
     If BtPlus <> Nil Then BtPlus.OnClick := BAjoutClick;

     BtMoins := TToolBarButton97(GetControl('BSUPP'));
     If BtMoins <> Nil Then BtMoins.OnClick := BSuppClick;

     // Actions glisser/déposer sur les grilles
     GrilleDepart.OnDragDrop := Depose_Objet;
     GrilleDepart.OnDragOver := OnDragOver;
     GrilleDest.OnDragDrop := Depose_Objet;
     GrilleDest.OnDragOver := OnDragOver;
     GrilleDepart.OnMouseDown:= OnMouseDown;
     GrilleDest.OnMouseDown:= OnMouseDown;

     // Mise en forme des grilles
     GrilleDepart.ColAligns[0] := taCenter;
     GrilleDepart.ColAligns[1] := TaLeftJustify;
     GrilleDest.ColAligns[0]   := TaCenter;
     GrilleDest.ColAligns[1]   := TaLeftJustify;

     //PT3
     If Action = 'CONSULTATION' Then
     Begin
        SetControlVisible('BVALIDER', False);
        SetControlVisible('BAJOUT',   False);
        SetControlVisible('BSUPP',    False);
     End;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /
Description .. : Fermeture de l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.OnClose ;
begin
  Inherited ;
  If TobDispo <> Nil then FreeAndNil(TobDispo);
  If TobAssoc <> Nil then FreeAndNil(TobAssoc);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Affichage des données des TOBs en grille
Mots clefs ... :
*****************************************************************}
Procedure TOF_MULSTAGEEMPLOI.AfficheGrille;
begin
     // Pour éviter les problèmes de rafraîchissement
     GrilleDepart.RowCount := 2;
     GrilleDest.RowCount   := 2;
     GrilleDepart.Cells[0,1] := ''; GrilleDepart.Cells[1,1] := '';
     GrilleDest.Cells[0,1] := '';   GrilleDest.Cells[1,1]   := '';

     TobDispo.PutGridDetail(GrilleDepart,False,False,'CC_CODE;CC_LIBELLE');
     TobAssoc.PutGridDetail(GrilleDest,  False,False,'CC_CODE;CC_LIBELLE');
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Bouton Ajouter
Mots clefs ... : 
*****************************************************************}
Procedure TOF_MULSTAGEEMPLOI.BAjoutClick(Sender : TObject);
begin
  Depose_Objet(GrilleDest, GrilleDepart, -1, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Bouton Retirer
Mots clefs ... : 
*****************************************************************}
Procedure TOF_MULSTAGEEMPLOI.BSuppClick(Sender : TObject);
begin
  Depose_Objet(GrilleDepart, GrilleDest, -1, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Fonction de déplacement d'un emploi d'une grille à une 
Suite ........ : autre
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.Depose_Objet(Destination, Origine : TObject; X,Y : Integer);
  var
  Row : Integer;
begin
  if not ((Destination is THGrid) and (Origine is THGrid)) then exit;
  if (Destination = Origine) then Exit;
  if (x = -1) and (y = -1) then
  begin
    for Row := (Origine as THGrid).RowCount-1 downto 1 do
    begin
      if ((Origine as THGrid).IsSelected(Row)) then
      begin
        if (Destination as THGrid) = GrilleDepart then
          AssocToDispo(Row-1);
        if (Destination as THGrid) = GrilleDest then
          DispoToAssoc(Row-1);
      end;
    end;
  end else begin
    if (Destination as THGrid) = GrilleDepart then
      AssocToDispo(DragRow-1);
    if (Destination as THGrid) = GrilleDest then
      DispoToAssoc(DragRow-1);
  end;
  (Origine as THGrid).ClearSelected;
  (Destination as THGrid).ClearSelected;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Détermine si le drag&drop est possible
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.OnDragOver (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := (Sender is THgrid);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Détermination de l'élément qui sera déposé
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.OnMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
 Col : Integer;
begin
  if (Sender is THGrid) then
  begin
    (Sender as THGrid).MouseToCell(X,Y,Col, DragRow);
    if not (ssCtrl	in Shift) then
      (Sender as THGrid).BeginDrag(False);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : Passage d'un élément de la grille Associé à la grille 
Suite ........ : Disponible
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.AssocToDispo(i: integer);
begin
  if TobAssoc.FillesCount(1) >= i+1 then
  begin
     TobAssoc.FindFirst(['CC_CODE'],[GrilleDest.Cells[0,i+1]],False).ChangeParent(TobDispo, -1);
     TobDispo.Detail.Sort('CC_CODE');
     AfficheGrille;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/09/2007
Modifié le ... :   /  /
Description .. : Passage d'un élément de la grille Disponible à la grille 
Suite ........ : Associé
Mots clefs ... : 
*****************************************************************}
procedure TOF_MULSTAGEEMPLOI.DispoToAssoc(i: integer);
begin
  if TobDispo.FillesCount(1) >= i+1 then
  begin
     TobDispo.FindFirst(['CC_CODE'],[GrilleDepart.Cells[0,i+1]],False).ChangeParent(TobAssoc, -1);
     AfficheGrille;
  end;
end;

{TOF_LISTESTAGEEMPLOI}
{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Initialisation de la fiche
Mots clefs ... :
*****************************************************************}
procedure TOF_LISTESTAGEEMPLOI.OnArgument (S : String ) ;
begin
  Inherited ;

	//PT1 - Début
	If PGBundleCatalogue then //PT4
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PST_PREDEFINI',False);
			SetControlText   ('PST_PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT2
	end
	else
	begin
		SetControlVisible('PST_PREDEFINI', False);
		SetControlVisible('TPST_PREDEFINI',False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	SetControlText('XX_WHERE', DossiersAInterroger(GetControlText('PST_PREDEFINI'),GetControlText('NODOSSIER'),'PST'));
	//PT1 - Fin

	TFSaisieList(Ecran).BCherche.Click;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 15/02/2008 / PT1
Modifié le ... :   /  /
Description .. : Chargement des données
Mots clefs ... :
*****************************************************************}
procedure TOF_LISTESTAGEEMPLOI.OnLoad ;
begin
  Inherited ;
	SetControlText('XX_WHERE', DossiersAInterroger(GetControlText('PST_PREDEFINI'),GetControlText('NODOSSIER'),'PST')); //PT1
end;

Initialization
  registerclasses ( [ TOF_MULSTAGEEMPLOI,TOF_LISTESTAGEEMPLOI ] ) ;
end.



