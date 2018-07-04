{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PROFILCOMPTPRES ()
Mots clefs ... : TOF;PROFILCOMPTPRES
*****************************************************************}
{
PT1  09/08/2007  GGU  Recalcul automatique des compteurs lors d'une modification ou d'une suppresion de compteur.
PT2  27/08/2007  FLO  Ajout de l'association inverse : compteur vers profils 
}
Unit UTofPG_PROFILCOMPTPRES;

Interface

Uses StdCtrls, Controls,  Classes,
{$IFNDEF EAGLCLIENT}
     db, utob, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,   Fe_Main,
{$else}
     eMul, uTob,  MaineAgl,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97 ;

Type
  TOF_PROFILCOMPTPRES = Class (TOF)
    procedure OnNew                   ; override ;
    procedure OnDelete                ; override ;
    procedure OnUpdate                ; override ;
    procedure OnLoad                  ; override ;
    procedure OnArgument (S : String) ; override ;
    procedure OnDisplay               ; override ;
    procedure OnClose                 ; override ;
    procedure OnCancel                ; override ;
  private
    Control_BTNASSOC, Control_BTNDISPO : TToolbarButton97;
    Control_LPROFIL : THLabel;
    Control_GRIDASSOC, Control_GRIDDISPO : THGrid;
    TobDispo, TobNonDispo, TobAssoc, TobPoubelle : TOB;
    DragRow : Integer;
    Modifie : Boolean;
    Profil  : String;
    DateValid : TDateTime;
    Compteur,Libelle : String; //PT2
    ModeProfil : Boolean; //PT2
    procedure Depose_Objet(Destination, Origine : TObject; X,Y : Integer);
    procedure AfficheGrids;
    procedure OnClickDispoToAssoc(Sender : TObject);
    procedure OnClickAssocToDispo(Sender : TObject);
    procedure AssocToDispo(i: integer);
    procedure DispoToAssoc(i: integer);
    procedure OnDragOver (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure OnMouseDown(Sender : TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

    procedure ChargeCompteursAssocies;   //PT2
    procedure ChargeProfilsAssocies;     //PT2
  end;

Implementation

uses
  Variants, Grids, ed_tools
  , PGPresence //PT1
  ; //, uTobDebug

procedure TOF_PROFILCOMPTPRES.AssocToDispo( i : integer);
var T : TOB;
begin
  if TobAssoc.FillesCount(1) >= i+1 then
  begin
     If ModeProfil Then //PT2
     Begin
          TobNonDispo.FindFirst(['PYR_COMPTEURPRES'],[Control_GRIDASSOC.Cells[0,i+1]],False).ChangeParent(TobDispo, -1);
          TobAssoc.FindFirst(['PPV_COMPTEURPRES'],[Control_GRIDASSOC.Cells[0,i+1]],False).ChangeParent(TobPoubelle, -1);
     End
     Else
     Begin
          TobNonDispo.FindFirst(['PPQ_PROFILPRES'],[Control_GRIDASSOC.Cells[0,i+1]],False).ChangeParent(TobDispo, -1);
          TobAssoc.FindFirst(['PPV_PROFILPRES'],[Control_GRIDASSOC.Cells[0,i+1]],False).ChangeParent(TobPoubelle, -1);
     End;
    AfficheGrids;
  end;
end;

procedure TOF_PROFILCOMPTPRES.DispoToAssoc( i : integer);
begin
  if TobDispo.FillesCount(1) >= i+1 then
  begin
    with TOB.Create ('PROFILCOMPTEURPRES', TobAssoc, -1) do
    begin
      If ModeProfil Then //PT2
      Begin
        PutValue ('PPV_PROFILPRES'  , Profil);
        PutValue ('PPV_COMPTEURPRES', TobDispo.Detail[i].GetString('PYR_COMPTEURPRES'));
        PutValue ('PPV_LIBELLE'     , TobDispo.Detail[i].GetString('PYR_LIBELLE'));
        PutValue ('PPV_DATEVALIDITE', DateValid);
      End
      Else
      Begin
        PutValue ('PPV_PROFILPRES'  , TobDispo.Detail[i].GetString('PPQ_PROFILPRES'));
        PutValue ('PPV_DATEVALIDITE', TobDispo.Detail[i].GetString('PPQ_DATEVALIDITE'));
        PutValue ('PPV_COMPTEURPRES', Compteur);
        PutValue ('PPV_LIBELLE'     , Libelle);
        AddChampSupValeur('LIBELLEPROFIL', TobDispo.Detail[i].GetString('PPQ_LIBELLE'));
      End;
    end;
    TobDispo.Detail[i].ChangeParent(TobNonDispo, -1);
    AfficheGrids;
  end;
end;

procedure TOF_PROFILCOMPTPRES.OnNew ;
begin
  Inherited ;
  Modifie := False;
end ;

procedure TOF_PROFILCOMPTPRES.OnDelete ;
begin
  Inherited ;
  CompteursARecalculer(DateValid); //PT1
end ;

procedure TOF_PROFILCOMPTPRES.OnUpdate ;
var
libellemodif,libellemodif2, daterecalcul, paramrecalcul : string;
begin
  Inherited ;
  BEGINTRANS;
  try
    If ModeProfil Then //PT2
    Begin
      ExecuteSQL('DELETE FROM PROFILCOMPTEURPRES WHERE PPV_PROFILPRES = "'+Profil+'" AND PPV_DATEVALIDITE = "'+USDATETIME(DateValid)+'"');
    End
    Else
    Begin
      ExecuteSQL('DELETE FROM PROFILCOMPTEURPRES WHERE PPV_COMPTEURPRES = "'+Compteur+'" AND PPV_LIBELLE = "'+Libelle+'"');
    End;
    TobAssoc.SetAllModifie(True);
    TobAssoc.InsertOrUpdateDB;
   // CompteursARecalculer(DateValid); //PT1

   // une fois la mise à jour effectuée : affichage et saisie de la date de recakcul du compteur
   if ModeProfil then
   begin
      If (PresenceProfileIsUsed(Profil)) Then
      begin
         Daterecalcul := datetostr(datevalid);
         libellemodif := 'Le profil de présence ' + Profil + ' ' + libelle;
         libellemodif2 := 'a été modifié.';
         paramrecalcul := libellemodif + ';' + libellemodif2 + ';' + daterecalcul;
         AglLanceFiche('PAY','PRESRECALCUL','','',paramrecalcul);
      end;
   end
   else
   begin
      If (PresenceMeterIsUsed(Compteur)) Then
      begin
          Daterecalcul := datetostr(datevalid);
          libellemodif := 'Le compteur de présence ' + compteur  + ' ' +  libelle;
          libellemodif2 := 'a été modifié';
          paramrecalcul := libellemodif + ';' + libellemodif2 + ';' + daterecalcul;
          AglLanceFiche('PAY','PRESRECALCUL','','',paramrecalcul);
      end;
   end;

  except
    ROLLBACK;
  end;
  COMMITTRANS;
  Modifie := False;
  Ecran.Close;
end ;

procedure TOF_PROFILCOMPTPRES.OnLoad ;
begin
  Inherited ;

  TobDispo := TOB.Create('Tob Dispo',nil,-1);
  TobAssoc := TOB.Create('Tob PROFILCOMPTEURPRES',nil,-1);
  TobNonDispo := TOB.Create('Tob Non Dispo',nil,-1);
  TobPoubelle := TOB.Create('Tob Poubelle',nil,-1);

  //PT2 - Début
  If ModeProfil Then
     ChargeCompteursAssocies
  Else
     ChargeProfilsAssocies;
  //PT2 - Fin

  AfficheGrids;
  Modifie := False;
end ;

procedure TOF_PROFILCOMPTPRES.OnArgument (S : String ) ;
begin
  Inherited ;
  Profil := Trim(ReadTokenPipe(S, ';'));
  //PT2 - Début
  If Profil = 'COMPTEUR' Then
  Begin
     Profil := '';
     Compteur := Trim(ReadTokenPipe(S, ';'));
     ModeProfil := False;
  End
  Else
     ModeProfil := True;
  //PT2 - Fin
  DateValid := AGLStrToDate(Trim(ReadTokenPipe(S, ';')));
  Libelle := Trim(ReadTokenPipe(S, ';')); //PT2

  Control_LPROFIL  := (GetControl('LPROFIL')   as THlabel);
  If ModeProfil Then //PT2
       Control_LPROFIL.Caption := Profil + ' - '+ Libelle
  Else
       Control_LPROFIL.Caption := Compteur + ' - '+ Libelle;

  Control_BTNASSOC := (GetControl('BTN_ASSOC') as TToolbarButton97);
  Control_BTNDISPO := (GetControl('BTN_DISPO') as TToolbarButton97);
  Control_GRIDASSOC:= (GetControl('GRID_ASSOC')as THGrid);
  Control_GRIDDISPO:= (GetControl('GRID_DISPO')as THGrid);
  Control_BTNASSOC.OnClick := OnClickDispoToAssoc;
  Control_BTNDISPO.OnClick := OnClickAssocToDispo;
  Control_GRIDASSOC.OnDragDrop := Depose_Objet;
  Control_GRIDASSOC.OnDragOver := OnDragOver;
  Control_GRIDDISPO.OnDragDrop := Depose_Objet;
  Control_GRIDDISPO.OnDragOver := OnDragOver;
  Control_GRIDASSOC.OnMouseDown:= OnMouseDown;
  Control_GRIDDISPO.OnMouseDown:= OnMouseDown;

  //PT2 - Début
  // Adaptation de l'écran au mode Compteur
  // Titre de l'écran
  If Not ModeProfil Then
  Begin
       Ecran.Caption := 'Profils faisant référence au compteur';
       UpdateCaption(Ecran);

       // Intitulé de la référence
       (GetControl('LBL_PROFIL') as THLabel).Caption := 'Compteur :';

       // Intitulés de listboxes
       (GetControl('LBL_DISPO') as THLabel).Caption := 'Profils disponibles';
       (GetControl('LBL_ASSOC') as THLabel).Caption := 'Profils associés';

       // Entête de colonne
       Control_GRIDDISPO.Cells[0,0] := 'Profil';
       Control_GRIDASSOC.Cells[0,0] := 'Profil';
  End;
  //PT2 - Fin
end ;

procedure TOF_PROFILCOMPTPRES.OnClose ;
begin
  FreeAndNil(TobDispo);
  FreeAndNil(TobAssoc);
  FreeAndNil(TobNonDispo);
  FreeAndNil(TobPoubelle);
  Inherited ;
end ;

procedure TOF_PROFILCOMPTPRES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PROFILCOMPTPRES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PROFILCOMPTPRES.OnClickAssocToDispo(Sender: TObject);
begin
  Depose_Objet(Control_GRIDDISPO, Control_GRIDASSOC, -1, -1);
end;

procedure TOF_PROFILCOMPTPRES.OnClickDispoToAssoc(Sender: TObject);
begin
  Depose_Objet(Control_GRIDASSOC, Control_GRIDDISPO, -1, -1);
end;

procedure TOF_PROFILCOMPTPRES.Depose_Objet(Destination, Origine : TObject; X,Y : Integer);
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
          AssocToDispo(Row-1);
        if (Destination as THGrid) = Control_GRIDASSOC then
          DispoToAssoc(Row-1);
      end;
    end;
  end else begin
    if (Destination as THGrid) = Control_GRIDDISPO then
      AssocToDispo(DragRow-1);
    if (Destination as THGrid) = Control_GRIDASSOC then
      DispoToAssoc(DragRow-1);
  end;
  (Origine as THGrid).ClearSelected;
  (Destination as THGrid).ClearSelected;
  AfficheGrids;
end;

procedure TOF_PROFILCOMPTPRES.OnDragOver (Sender, Source: TObject; X, Y: Integer;
                                          State: TDragState; var Accept: Boolean);
begin
  Accept := (Sender is THgrid);
end;

procedure TOF_PROFILCOMPTPRES.AfficheGrids;
var
  i, j : integer;
begin
  for i := 1 to Control_GRIDASSOC.RowCount -1 do
  begin
    for j := 0 to Control_GRIDASSOC.ColCount -1 do
    begin
      Control_GRIDASSOC.Cells[j,i] := '';
      Control_GRIDDISPO.CellValues[j,i] := '';
    end;
  end;
  Control_GRIDASSOC.RowCount := 2;
  Control_GRIDDISPO.RowCount := 2;

  If ModeProfil Then //PT2
  Begin
       TobAssoc.PutGridDetail(Control_GRIDASSOC,False,True,'PPV_COMPTEURPRES;PPV_LIBELLE');
       TobDispo.PutGridDetail(Control_GRIDDISPO,False,True,'PYR_COMPTEURPRES;PYR_LIBELLE');
  End
  Else
  Begin
       TobAssoc.PutGridDetail(Control_GRIDASSOC,False,True,'PPV_PROFILPRES;LIBELLEPROFIL');
       TobDispo.PutGridDetail(Control_GRIDDISPO,False,True,'PPQ_PROFILPRES;PPQ_LIBELLE');
  End;
end;

procedure TOF_PROFILCOMPTPRES.OnMouseDown(Sender: TObject;
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
Auteur  ...... : GGU
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. : Chargement des compteurs pour un profil
Mots clefs ... : 
*****************************************************************}
procedure TOF_PROFILCOMPTPRES.ChargeCompteursAssocies;
var
  i : integer;
  Q : TQuery;
  stQ : String;
begin
  TobDispo.LoadDetailFromSQL(' SELECT PYR_COMPTEURPRES, PYR_LIBELLE, PYR_DATEVALIDITE '
                            +' FROM COMPTEURPRESENCE a'
                            +' WHERE ##PYR_PREDEFINI## '
                            +' and a.PYR_DATEVALIDITE  = (select max(b.PYR_DATEVALIDITE )'
                            +'                              from COMPTEURPRESENCE b'
                            +'                             where a.PYR_COMPTEURPRES = b.PYR_COMPTEURPRES)'
                          //  +'                               AND b.PYR_DATEVALIDITE <= "'+USDATETIME(DateValid)+'" )'
                            +' ORDER BY PYR_COMPTEURPRES');
  //On supprime de la liste des dispo ceux qui sont utilisés
  stQ   := 'SELECT PPV_PROFILPRES, PPV_COMPTEURPRES, PPV_LIBELLE'
         + ' FROM PROFILCOMPTEURPRES '
         + 'WHERE PPV_PROFILPRES = "'+Profil+'" '
         + '  AND PPV_DATEVALIDITE = "'+USDATETIME(DateValid)+'"';
  Q := OpenSQL(stQ,True);
  while not Q.Eof do
  begin
    //On parcours la liste des Dispo pour retrouver ceux qui sont associés
    for i := 0 to TobDispo.FillesCount(1)-1 do
    begin
      if TobDispo.Detail[i].GetString('PYR_COMPTEURPRES') = Q.FindField('PPV_COMPTEURPRES').AsString then
      begin
        DispoToAssoc(i);
        break;
      end
    end;
    Q.Next;
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 27/08/2007 / PT2
Modifié le ... :   /  /    
Description .. : Chargement des profils pour un compteur
Mots clefs ... : 
*****************************************************************}
procedure TOF_PROFILCOMPTPRES.ChargeProfilsAssocies;
var
  i : integer;
  T : TOB;
  Q : TQuery;
begin
     // Chargement de tous les profils d'abord
     TobDispo.LoadDetailFromSQL(' SELECT PPQ_PROFILPRES, PPQ_DATEVALIDITE, PPQ_LIBELLE'
                            +' FROM PROFILPRESENCE PP'
                            +' WHERE PP.PPQ_DATEVALIDITE  = (SELECT MAX(PPM.PPQ_DATEVALIDITE)'
                                                          +' FROM PROFILPRESENCE PPM'
                                                          +' WHERE PP.PPQ_PROFILPRES = PPM.PPQ_PROFILPRES)'
                                                    //      +' AND PPM.PPQ_DATEVALIDITE <= "'+UsDateTime(DateValid)+'")'
                            +' ORDER BY PPQ_PROFILPRES');

     // Chargement des profils rattachés au compteur
     Q := OpenSQL('SELECT PPV_PROFILPRES,PPV_DATEVALIDITE,PPV_COMPTEURPRES,PPV_LIBELLE '+
                              'FROM PROFILCOMPTEURPRES P1 ' +
                              'WHERE PPV_COMPTEURPRES="'+Compteur+'" AND '+
                              'PPV_DATEVALIDITE=(SELECT MAX (P2.PPV_DATEVALIDITE) FROM PROFILCOMPTEURPRES P2 '+
                              'WHERE P2.PPV_PROFILPRES=P1.PPV_PROFILPRES )'+
                        //      'AND P2.PPV_DATEVALIDITE>="'+UsDateTime(DateValid)+'")'+
                              'ORDER BY PPV_PROFILPRES',True);
     while not Q.Eof do
     begin
          //On parcours la liste des Dispo pour retrouver ceux qui sont associés
          for i := 0 to TobDispo.FillesCount(1)-1 do
          begin
               if TobDispo.Detail[i].GetString('PPQ_PROFILPRES') = Q.FindField('PPV_PROFILPRES').AsString then
               begin
                    DispoToAssoc(i);
                    break;
               end
          end;
          Q.Next;
     end;
     Ferme(Q);
end;

Initialization
  registerclasses ( [ TOF_PROFILCOMPTPRES ] ) ;
end.
