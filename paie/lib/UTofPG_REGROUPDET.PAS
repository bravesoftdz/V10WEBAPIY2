{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 25/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGREGROUPDET ()
Mots clefs ... : TOF;PGREGROUPDET
*****************************************************************}

Unit UTofPG_REGROUPDET;

Interface

Uses StdCtrls, Controls,  Classes,
{$IFNDEF EAGLCLIENT}
     db, utob, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, 
{$else}
     eMul, uTob,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97 ;

Type
  TOF_REGROUPDET = Class (TOF)
    procedure OnNew                   ; override ;
    procedure OnDelete                ; override ;
    procedure OnUpdate                ; override ;
    procedure OnLoad                  ; override ;
    procedure OnArgument (S : String) ; override ;
    procedure OnDisplay               ; override ;
    procedure OnClose                 ; override ;
    procedure OnCancel                ; override ;
  private
    Control_BTNASSOCPLUS, Control_BTNDISPOPLUS, Control_BTNASSOCMOINS, Control_BTNDISPOMOINS,
    Control_BTNMOINSPLUS, Control_BTNPLUSMOINS: TToolbarButton97;
    Control_LREGROUPEMENT : THLabel;
    Control_GRIDASSOCAJOUT,Control_GRIDASSOCSOUST, Control_GRIDDISPO : THGrid;
    TobDispo, TobNonDispo, TobAssocAjout,TobAssocSoust, TobPoubelle : TOB;
    DragRow : Integer;
    Modifie : Boolean;
    codeassoc,Libelle, pred, dos, Typeregroupement ,Regroupement ,Themeregroupement: String;

    procedure Depose_Objet(Destination, Origine : TObject; X,Y : Integer);
    procedure AfficheGrids;
    procedure OnClickDispoToAssocAjout(Sender : TObject);
    procedure OnClickAjoutSoust(Sender : TObject);
    procedure OnClickSoustAjout(Sender : TObject);
    procedure OnClickAssocToDispoAjout(Sender : TObject);
    procedure OnClickDispoToAssocSoust(Sender : TObject);
    procedure OnClickAssocToDispoSoust(Sender : TObject);
    procedure AssocToDispo(sens: string;i: integer);
    procedure DispoToAssoc(sens : string; i: integer);
    procedure Ajoutsoust(sens : string; i: integer);
    procedure OnDragOver (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure OnMouseDown(Sender : TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

    procedure ChargeEvtAssocies;

  end;

Implementation

uses
  Variants, Grids, ed_tools, HSysMenu;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /
Description .. : Déplacement ASSOCIATION   vers DISPO
Mots clefs ... :
*****************************************************************}
procedure TOF_REGROUPDET.AssocToDispo( sens : string; i : integer);
begin
 if sens = '+' then
 begin
    if TobAssocAjout.FillesCount(1) >= i+1 then
    begin
       TobNonDispo.FindFirst(['CODE'],[Control_GRIDASSOCAJOUT.Cells[0,i+1]],False).ChangeParent(TobDispo, -1);
       TobAssocAJOUT.FindFirst(['POT_PGELEMENT'],[Control_GRIDASSOCAJOUT.Cells[0,i+1]],False).ChangeParent(TobPoubelle, -1);
    End;

  end
  else
  begin
    if TobAssocSoust.FillesCount(1) >= i+1 then
    begin
       TobNonDispo.FindFirst(['CODE'],[Control_GRIDASSOCSoust.Cells[0,i+1]],False).ChangeParent(TobDispo, -1);
       TobAssocSoust.FindFirst(['POT_PGELEMENT'],[Control_GRIDASSOCSoust.Cells[0,i+1]],False).ChangeParent(TobPoubelle, -1);
    End;
  end;
    AfficheGrids;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Déplacement DISPO vers ASSOCIATION
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.DispoToAssoc(sens: string; i : integer);

begin

  if TobDispo.FillesCount(1) >= i+1 then
  begin
        if sens = '+' then
        begin
        with TOB.Create ('PGREGROUPDET', tobassocAjout, -1) do
        begin
          PutValue ('POT_PREDEFINI'  , pred);
          PutValue ('POT_NODOSSIER'  , dos);
          PutValue ('POT_PGELEMENT', TobDispo.Detail[i].GetString('CODE'));
          PutValue ('POT_CODEREGROUP', Regroupement);
          PutValue ('POT_PLUSMOINS', '+');
          AddChampSupValeur('LIBELLE', TobDispo.Detail[i].GetString('LIBELLE'));
       end;
       end
       else
       begin
        with TOB.Create ('PGREGROUPDET', tobassocSoust, -1) do
        begin
          PutValue ('POT_PREDEFINI'  , pred);
          PutValue ('POT_NODOSSIER'  , dos);
          PutValue ('POT_PGELEMENT', TobDispo.Detail[i].GetString('CODE'));
          PutValue ('POT_CODEREGROUP', Regroupement);
          PutValue ('POT_PLUSMOINS', '-');
          AddChampSupValeur('LIBELLE', TobDispo.Detail[i].GetString('LIBELLE'));
        end;
      end;
  end;
    TobDispo.Detail[i].ChangeParent(TobNonDispo, -1);


    AfficheGrids;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Déplace objet de association Ajout vers SOUST ou 
Suite ........ :                objet de association SOUST vers Ajout
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.Ajoutsoust( sens : string; i : integer);
begin
 if sens = '+' then
 begin
    with TOB.Create ('PGREGROUPDET', tobassocAjout, -1) do
        begin
          PutValue ('POT_PREDEFINI'  , pred);
          PutValue ('POT_NODOSSIER'  , dos);
          PutValue ('POT_PGELEMENT', TobAssocSoust.Detail[i].GetString('POT_PGELEMENT'));
          PutValue ('POT_CODEREGROUP', Regroupement);
          PutValue ('POT_PLUSMOINS', '+');
          AddChampSupValeur('LIBELLE', TobAssocSoust.Detail[i].GetString('LIBELLE'));
        end;

    TobAssocSoust.Detail[i].ChangeParent(Tobpoubelle, -1);
 end
 else
 begin
      with TOB.Create ('PGREGROUPDET', tobassocSoust, -1) do
        begin
          PutValue ('POT_PREDEFINI'  , pred);
          PutValue ('POT_NODOSSIER'  , dos);
          PutValue ('POT_PGELEMENT', TobAssocAjout.Detail[i].GetString('POT_PGELEMENT'));
          PutValue ('POT_CODEREGROUP', Regroupement);
          PutValue ('POT_PLUSMOINS', '-');
          AddChampSupValeur('LIBELLE', TobAssocAjout.Detail[i].GetString('LIBELLE'));
        end;

    TobAssocAjout.Detail[i].ChangeParent(Tobpoubelle, -1);
 end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Création
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.OnNew ;
begin
  Inherited ;
  Modifie := False;
end ;

procedure TOF_REGROUPDET.OnDelete ;
begin
  Inherited ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.OnUpdate ;
begin
  Inherited ;
  BEGINTRANS;
  try
    Begin
      ExecuteSQL('DELETE FROM PGREGROUPDET WHERE POT_CODEREGROUP = "'+Regroupement+'"');
    End;
    TobAssocAjout.SetAllModifie(True);
    TobAssocSoust.SetAllModifie(True);
    TobAssocAjout.InsertOrUpdateDB;
    TobAssocSoust.InsertOrUpdateDB;
  except
    ROLLBACK;
  end;
  COMMITTRANS;
  Modifie := False;
  Ecran.Close;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Chargement
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.OnLoad ;
begin
  Inherited ;

  TobDispo := TOB.Create('Tob Dispo',nil,-1);
  TobAssocAjout := TOB.Create('Tob ASSOCIATION AJOUT',nil,-1);
  TobAssocSoust := TOB.Create('Tob ASSOCIATION SOUSTRACTION',nil,-1);
  TobNonDispo := TOB.Create('Tob Non Dispo',nil,-1);
  TobPoubelle := TOB.Create('Tob Poubelle',nil,-1);


  If typeregroupement = 'MAB'  Then
     ChargeEVTAssocies;


  AfficheGrids;
  Modifie := False;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 25/09/2007
Modifié le ... :   /  /    
Description .. : on argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.OnArgument (S : String ) ;
begin
  Inherited ;
  Typeregroupement   := Trim(ReadTokenPipe(S, ';'));
  Themeregroupement :=  Trim(ReadTokenPipe(S, ';'));
  Regroupement := Trim(ReadTokenPipe(S, ';'));

  Libelle := Trim(ReadTokenPipe(S, ';'));
  Pred:= Trim(ReadTokenPipe(S, ';'));
  Dos := Trim(ReadTokenPipe(S, ';'));

  Control_LREGROUPEMENT := (GetControl('LREGROUPEMENT')   as THlabel);
  Control_LREGROUPEMENT.Caption := Regroupement + ' - '+ Libelle ;


  Control_BTNASSOCPLUS := (GetControl('BTN_ASSOCPLUS') as TToolbarButton97);
  Control_BTNASSOCMOINS := (GetControl('BTN_ASSOCMOINS') as TToolbarButton97);
  Control_BTNDISPOPLUS := (GetControl('BTN_DISPOPLUS') as TToolbarButton97);
  Control_BTNDISPOMOINS := (GetControl('BTN_DISPOMOINS') as TToolbarButton97);
  Control_BTNMOINSPLUS := (GetControl('BTN_MOINSPLUS') as TToolbarButton97);
  Control_BTNPLUSMOINS := (GetControl('BTN_PLUSMOINS') as TToolbarButton97);
  Control_GRIDASSOCAJOUT:= (GetControl('GRID_ASSOCAJOUT')as THGrid);
  Control_GRIDASSOCSOUST:= (GetControl('GRID_ASSOCSOUST')as THGrid);
  Control_GRIDDISPO:= (GetControl('GRID_DISPO')as THGrid);
  Control_BTNASSOCPLUS.OnClick := OnClickDispoToAssocAjout;
  Control_BTNASSOCMOINS.OnClick := OnClickDispoToAssocSoust;
  Control_BTNDISPOPlus.OnClick := OnClickAssocToDispoAjout;
  Control_BTNDISPOMoins.OnClick := OnClickAssocToDispoSoust;
  Control_BTNPLUSMOINS.OnClick := OnClickAjoutSoust;
  Control_BTNMOINSPLUS.OnClick := OnClickSoustAjout;
  Control_GRIDASSOCAjout.OnDragDrop := Depose_Objet;
  Control_GRIDASSOCSoust.OnDragDrop := Depose_Objet;
  Control_GRIDASSOCAjout.OnDragOver := OnDragOver;
  Control_GRIDASSOCSoust.OnDragOver := OnDragOver;
  Control_GRIDDISPO.OnDragDrop := Depose_Objet;
  Control_GRIDDISPO.OnDragOver := OnDragOver;
  Control_GRIDASSOCAjout.OnMouseDown:= OnMouseDown;
  Control_GRIDASSOCSoust.OnMouseDown:= OnMouseDown;
  Control_GRIDDISPO.OnMouseDown:= OnMouseDown;


end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Ferme fiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.OnClose ;
begin
  FreeAndNil(TobDispo);
  FreeAndNil(TobAssocajout);
  FreeAndNil(TobAssocSoust);
  FreeAndNil(TobNonDispo);
  FreeAndNil(TobPoubelle);
  Inherited ;
end ;

procedure TOF_REGROUPDET.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_REGROUPDET.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /
Description .. : Déplace élèment ASSOCIATION AJOUT vers DISPO
Mots clefs ... :
*****************************************************************}
procedure TOF_REGROUPDET.OnClickAssocToDispoAjout(Sender: TObject);
begin
  Depose_Objet(Control_GRIDDISPO, Control_GRIDASSOCAJOUT, -1, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /
Description .. : Déplace élèment ASSOCIATION SOUSTRAIT vers DISPO
Mots clefs ... :
*****************************************************************}
procedure TOF_REGROUPDET.OnClickAssocToDispoSOUST(Sender: TObject);
begin
  Depose_Objet(Control_GRIDDISPO, Control_GRIDASSOCSOUST, -1, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /
Description .. : Déplace élément de DISPO vers ASSOCIATION AJOUT
Mots clefs ... :
*****************************************************************}
procedure TOF_REGROUPDET.OnClickDispoToAssocAjout(Sender: TObject);
begin
  Depose_Objet(Control_GRIDASSOCAjout, Control_GRIDDISPO, -1, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /
Description .. : Déplace élément de DISPO vers ASSOCIATION SOUSTRAIT
Mots clefs ... :
*****************************************************************}
procedure TOF_REGROUPDET.OnClickDispoToAssocSoust(Sender: TObject);
begin
  Depose_Objet(Control_GRIDASSOCSoust, Control_GRIDDISPO, -1, -1);
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /
Description .. : Déplace élément de AJOUT à SOUSTRACTION
Mots clefs ... :
*****************************************************************}
procedure TOF_REGROUPDET.OnClickAJOUTSOUST(Sender: TObject);
begin
  Depose_Objet(Control_GRIDASSOCSOUST, Control_GRIDASSOCAJOUT, -1, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /
Description .. : Déplace élément de SOUSTRACTION à  AJOUT
Mots clefs ... :
*****************************************************************}
procedure TOF_REGROUPDET.OnClickSOUSTAJOUT(Sender: TObject);
begin
  Depose_Objet(Control_GRIDASSOCAjout, Control_GRIDASSOCSoust, -1, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Dépose objet
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.Depose_Objet(Destination, Origine : TObject; X,Y : Integer);
  var
  Row : Integer;
begin
  if not ((Destination is THGrid) and (Origine is THGrid)) then exit;
  if (Destination = Origine) then Exit;
  Modifie:= TRUE;
  if (x = -1) and (y = -1) then
  begin
    for Row := (Origine as THGrid).RowCount-1 downto 1 do
    begin
      if ((Origine as THGrid).IsSelected(Row)) then
      begin
        if (Destination as THGrid) = Control_GRIDDISPO then
          if (Origine as THGRID) = Control_GRIDASSOCAJOUT then AssocToDispo('+',Row-1) else AssocToDispo('-',Row-1) ;
        if (Destination as THGrid) = Control_GRIDASSOCAjout then
          if (origine as THGrid) = Control_GRIDDISPO then  DispoToAssoc('+',Row-1) else ajoutsoust('+',Row-1);
        if (Destination as THGrid) = Control_GRIDASSOCSoust then
          if (origine as THGrid) = Control_GRIDDISPO then DispoToAssoc('-',Row-1) else ajoutsoust('-',Row-1);

      end;
    end;
  end else begin
    if (Destination as THGrid) = Control_GRIDDISPO then
     if (Origine as THGRID) = Control_GRIDASSOCAJOUT then AssocToDispo('+',DragRow-1) else AssocToDispo('-',DragRow-1) ;
     if (Destination as THGrid) = Control_GRIDASSOCAjout then
      if (origine as THGrid) = Control_GRIDDISPO then DispoToAssoc('+',DragRow-1)else ajoutsoust('+',Row-1);
     if (Destination as THGrid) = Control_GRIDASSOCsoust then
      if (origine as THGrid) = Control_GRIDDISPO then DispoToAssoc('-',DragRow-1)else ajoutsoust('-',Row-1);
  end;
  (Origine as THGrid).ClearSelected;
  (Destination as THGrid).ClearSelected;
  AfficheGrids;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : On drag over
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.OnDragOver (Sender, Source: TObject; X, Y: Integer;
                                          State: TDragState; var Accept: Boolean);
begin
  Accept := (Sender is THgrid);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Affichage de la grille
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.AfficheGrids;
var
HMTrad: THSystemMenu;
i : integer;
begin
       HMTrad.ResizeGridColumns(Control_GRIDASSOCAJOUT);
       HMTrad.ResizeGridColumns(Control_GRIDASSOCSOUST);
       HMTrad.ResizeGridColumns(Control_GRIDDISPO);

       for i := 1 to Control_GRIDASSOCAJOUT.RowCount -1 do
       begin
          Control_GRIDASSOCAJOUT.CellValues[0,i] := '';
          Control_GRIDASSOCAJOUT.CellValues[1,i] := '';
       end;

       for i := 1 to Control_GRIDASSOCSOUST.RowCount -1 do
       begin
          Control_GRIDASSOCSOUST.CellValues[0,i] := '';
          Control_GRIDASSOCSOUST.CellValues[1,i] := '';
       end;

       for i := 1 to Control_GRIDASSOCSOUST.RowCount -1 do
       begin
          Control_GRIDDISPO.CellValues[0,i] := '';
          Control_GRIDDISPO.CellValues[1,i] := '';
       end;

       Control_GRIDASSOCAJOUT.RowCount := 2;
       Control_GRIDDISPO.RowCount := 2;
       Control_GRIDASSOCSOUST.RowCount := 2;

       TobAssocSoust.PutGridDetail(Control_GRIDASSOCSOUST,False,True,'POT_PGELEMENT;LIBELLE');
       TobAssocAjout.PutGridDetail(Control_GRIDASSOCAJOUT,False,True,'POT_PGELEMENT;LIBELLE');
       TobDispo.PutGridDetail(Control_GRIDDISPO,False,True,'CODE;LIBELLE');

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 27/09/2007
Modifié le ... :   /  /    
Description .. : Déplacement avec la souris
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.OnMouseDown(Sender: TObject;
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
Auteur  ...... : NA
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. : Chargement des motifs évènement associés
Mots clefs ... : 
*****************************************************************}
procedure TOF_REGROUPDET.ChargeEVTAssocies;
var
  i : integer;
  Q : TQuery;
  stQ : String;
begin

  if themeregroupement = '010' // quantité
  then
  TobDispo.LoadDetailFromSQL(' SELECT PMA_MOTIFABSENCE as CODE, PMA_LIBELLE as LIBELLE from MOTIFABSENCE WHERE ##PMA_PREDEFINI## '+
  ' PMA_JOURHEURE = "QTE" ORDER BY PMA_MOTIFABSENCE')
  else
  begin
   if themeregroupement = '008'  // jour
   then
    TobDispo.LoadDetailFromSQL(' SELECT PMA_MOTIFABSENCE as CODE, PMA_LIBELLE as LIBELLE from MOTIFABSENCE WHERE ##PMA_PREDEFINI## '+
  ' PMA_JOURHEURE = "JOU" ORDER BY PMA_MOTIFABSENCE')
   else // jour et heure
   TobDispo.LoadDetailFromSQL(' SELECT PMA_MOTIFABSENCE as CODE, PMA_LIBELLE as LIBELLE from MOTIFABSENCE WHERE ##PMA_PREDEFINI## '+
  ' PMA_JOURHEURE <> "QTE" ORDER BY PMA_MOTIFABSENCE') ;
  end;


  //On supprime de la liste des dispo ceux qui sont utilisés
  stQ   := 'SELECT POT_PGELEMENT, POT_PLUSMOINS FROM PGREGROUPDET WHERE ##POT_PREDEFINI## POT_CODEREGROUP = "'+Regroupement+'" ' ;

  Q := OpenSQL(stQ,True);
  while not Q.Eof do
  begin
    //On parcours la liste des Dispo pour retrouver ceux qui sont associés
    for i := 0 to TobDispo.FillesCount(1)-1 do
    begin
      if TobDispo.Detail[i].GetString('CODE') = Q.FindField('POT_PGELEMENT').AsString then
      begin
        DispoToAssoc(Q.FindField('POT_PLUSMOINS').AsString,i);
        break;
      end
    end;
    Q.Next;
  end;
  Ferme(Q);
end;


Initialization
  registerclasses ( [ TOF_REGROUPDET ] ) ;
end.
