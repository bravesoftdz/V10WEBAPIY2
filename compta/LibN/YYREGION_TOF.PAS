{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYREGION ()
Mots clefs ... : TOF;YYREGION
*****************************************************************}
Unit YYREGION_TOF ;

Interface

Uses StdCtrls,
     Controls, 
     Classes,
     windows,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables, 
     mul,
{$else}
     eMul, 
     uTob,
{$ENDIF}
     forms,
     sysutils,                           
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     utob ;

Type
  TOF_YYREGION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private
        {$IFDEF EAGLCLIENT}
        procedure BNextQClick(Sender: TObject);
        procedure BPrevQClick(Sender: TObject);
        {$ENDIF}
        procedure Imprimer(Sender : TObject) ;
        procedure FListeRowEnter(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean) ;
        procedure FListeRowExit(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean) ;
        procedure FListeKeyPress( Sender : TObject ; var Key : Char ) ;
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure ParcoursListe(Sender: TObject);
        procedure ParametrageFiche;
        procedure ParametrageGrille(GRILLE: THGrid);
        procedure LancerRecherche(Pays:string);
        procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        function  EstRemplit(Ligne: Integer): Boolean;
        procedure RafraichirGrille(TOBLigne : TOB);
        function  VerifLigne(Fliste:THGrid):boolean;
        Function  ExisteEnreg(Fliste:THGrid;Ligne:integer):boolean;
        Function  EnregLigne(Sender: TOBject):boolean;
        procedure BoutonsNavMaj;
        Procedure SauvGrille;
        function  Supprim:boolean;
    public
        FListe          : THGrid ;
        RG_PAYS         : THValComboBox;
        Pays            : String;
        NatureEtat,CodeEtat : String;
        Action,
        NomTable,
        TitreEcran,
        ListeChamps,
        ListeColsUniq ,
        ChampTri        : String;
        Modif                   : boolean;
        DernLigne,FormKdown : integer;
        TOBLigne,
        TOBSauv         : TOB;
        procedure BInsertClick        (Sender : TObject);
        procedure BDeleteClick        (Sender : TObject);
        procedure BDefaireClick       (Sender : TObject);
        procedure OnChangeRG_PAYS     (Sender : TObject);
        procedure OnClickRG_PAYS      (Sender : TObject);
        procedure OnLoadRecord        (Sender:TObject)            ;
        function  OnUpdateRecord      (Sender:TObject)     : Integer   ;
        function  OnAfterUpdateRecord (Sender:TObject)     : Integer   ;
        function  OnDeleteRecord      (Sender:TObject)     : Integer   ;
        PROCEDURE BValider            (Sender: TOBject);
  end ;

Implementation

USES ut,printdbg;

procedure TOF_YYREGION.OnNew ;
begin
  Inherited ;
  pgiinfo('OnNew');
end ;

procedure TOF_YYREGION.OnDelete ;
begin
  Inherited ;
  pgiinfo('OnDelete');
end ;

procedure TOF_YYREGION.OnUpdate ;
begin
  Inherited ;
  pgiinfo('ONupdate');
end ;

procedure TOF_YYREGION.OnLoad ;
begin
  Inherited ;
  pgiinfo('OnLoad');
  Modif := False;
  Action:='';
  DernLigne := 1;
end ;

procedure TOF_YYREGION.OnArgument (S : String ) ;
begin
  Inherited ;
  FListe  := THGrid(GetControl('Grille', True));
  TButton(GetControl('BFirst',True)).onClick       := ParcoursListe;
  TButton(GetControl('BPrev',True)).onClick        := ParcoursListe;
  TButton(GetControl('BNext',True)).onClick        := ParcoursListe;
  TButton(GetControl('BLast',True)).onClick        := ParcoursListe;
  TButton(GetControl('BImprimer',True)).onClick    := Imprimer;
  TButton(GetControl('BInsert',True)).onClick      := BInsertClick;
  TButton(GetControl('BDelete',True)).onClick      := BDeleteClick;
  TButton(GetControl('bDefaire',True)).onClick     := BDefaireClick;
  TButton(GetControl('BValider',True)).onClick     := BValider;
  FListe.OnRowEnter                                := FListeRowEnter;
  FListe.OnRowExit                                 := FListeRowExit;
  FListe.OnCellEnter                               := GrilleCellEnter ;
  FListe.OnKeyPress                                := FListeKeyPress ;
  Ecran.OnKeyDown                                  := FormKeyDown;
  ThValComboBox(getcontrol('RG_PAYS',true)).OnClick := OnClickRG_PAYS;
  ThValComboBox(getcontrol('RG_PAYS',true)).onchange := OnChangeRG_PAYS;
{$IFDEF EAGLCLIENT}
  TButton(GetControl('BPrevQ',True)).onClick       := BPrevQClick ;
  TButton(GetControl('BNextQ',True)).onClick       := BNextQClick ;
  SetControlVisible('BPrevQ', True) ;
  SetControlVisible('BNextQ', True) ;
{$ENDIF}

  // paramètrage dans la fille :
  ParametrageFiche ;
  ParametrageGrille ( FListe ) ;
  TOBSauv :=TOB.Create('TOB sauvegarde',nil,-1);

  // Titre de l'écran
  Ecran.Caption := 'REGIONS' ;
  UpdateCaption( Ecran ) ;

  // Bouton imprimer
  SetControlVisible ('BImprimer', ( Trim(NatureEtat)<>'' ) and ( Trim(CodeEtat)<>'' ) ) ;

  Pays := TrouveArgument(S,'PAYS','') ;
  SetControlProperty('RG_PAYS', 'DATATYPE', 'TTPAYS') ;
  RG_PAYS := THValComboBox(GetControl('RG_PAYS',True));
  if Pays <> '' then
    begin
    SetControlText('RG_PAYS',Pays);
    RG_PAYS.Enabled := False;
    end
    else
    begin
    RG_PAYS.Items.Add('<<TOUS>>');
    RG_PAYS.ItemIndex := 0 ;
    end;
   LancerRecherche(RG_PAYS.TEXT);
end ;

procedure TOF_YYREGION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_YYREGION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYREGION.OnCancel () ;
begin
  Inherited ;
end ;


function TOF_YYREGION.EstRemplit(Ligne: Integer): Boolean;
var Col    : Integer ;
    Deb  : Integer ;
begin
  Result := False ;
  if FListe.DBIndicator then Deb := 1  else Deb := 0 ;
  for Col := Deb to FListe.RowCount - 1 do
     begin
     if Trim( FListe.Cells[col , Ligne] ) <> '' then
       begin
       Result := True ;
       Exit ;
       end ;
     end;
end;

procedure TOF_YYREGION.BInsertClick (Sender : TObject);
begin
   Action := 'INSERT';
   TOBSauv.AddChampSupValeur('PAYS',RG_PAYS.text);
   TOBSauv.GetGridDetail(Fliste,Fliste.Rowcount,NomTable,'RG_PAYS;RG_REGION;RG_LIBELLE',true);
   if RG_PAYS.Text <>'<<Tous>>' then
      begin
      FListe.InsertRow(1);
      FListe.GotoRow(1);
      FListe.SetFocus;
      end
      else
      PgiInfo('Pour saisir une nouvelle région, veuillez sélectionner un pays.','REGION');
end;

procedure TOF_YYREGION.BDeleteClick (Sender : TObject);
begin
   Action := 'DELETE';
   TOBSauv.ClearDetail;
   TOBSauv.AddChampSupValeur('PAYS',RG_PAYS.text);
   TOBSauv.GetGridDetail(Fliste,Fliste.Rowcount,NomTable,'RG_PAYS;RG_REGION;RG_LIBELLE',true);
   if Supprim and not ExisteSQL('SELECT * FROM REGION WHERE RG_PAYS="' + RG_PAYS.Text + '" '
                + 'AND RG_REGION<>"' + FListe.Cells[2,Fliste.Row] + '"' )
      then ExecuteSql('UPDATE PAYS SET PY_REGION="-" WHERE PY_PAYS="' + RG_PAYS.Text + '"') ;

   LancerRecherche(RG_PAYS.Values[RG_PAYS.ItemIndex]);
end;

procedure TOF_YYREGION.BDefaireClick(Sender : TObject);
begin
  Action := 'DEFAIRE';
end;

procedure TOF_YYREGION.OnChangeRG_PAYS(Sender : TObject);
begin
 inherited;
 LancerRecherche(RG_PAYS.Values[RG_PAYS.ItemIndex]);
end;

procedure TOF_YYREGION.OnClickRG_PAYS(Sender : TObject);
begin
    if (Action='INSERT') or (Action='MODIF') then
       begin
       if PGIAskCancel('Voulez-vous enregistrer les modifications ?',Ecran.Caption)=mrYes then
         EnregLigne(sender);
       end;
end;

procedure TOF_YYREGION.OnLoadRecord;
begin
 inherited;
   PgiInfo('OnLoadRecord');
end;

function  TOF_YYREGION.OnUpdateRecord(Sender:TObject)  : Integer   ;
var RG_REGION,RG_LIBELLE,Sql : string;
begin
  Result := 0 ;
  inherited;
  RG_REGION:= FListe.cells[1,DernLigne];
  RG_LIBELLE:= FListe.cells[2,DernLigne];
  if Action='INSERT' then
     Sql:='INSERT INTO Region(RG_PAYS,RG_REGION,RG_LIBELLE) values("'+ RG_PAYS.Values[RG_PAYS.ItemIndex] +'","'+RG_REGION+'","'+RG_LIBELLE+'")'
     else
     Sql:='UPDATE Region set RG_LIBELLE="'+RG_LIBELLE+'"';
  ExecuteSql(Sql);
  Action:='';
end;

function  TOF_YYREGION.OnAfterUpdateRecord (Sender:TObject) : Integer   ;
begin
  inherited;
  Pgibox('OnAfterUpdateRecord');
  Result := 0 ;
  try
    if not ExisteSQL('SELECT * FROM REGION WHERE RG_PAYS="' + RG_PAYS.Text + '" AND RG_LIBELLE<>"' + FListe.Cells[2,Fliste.Row] + '"' )
      then ExecuteSql('UPDATE PAYS SET PY_REGION="X" WHERE PY_PAYS="' + RG_PAYS.Text + '"') ;
  except
    Result := 1 ;
  end ;
end;

function  TOF_YYREGION.OnDeleteRecord (Sender:TObject)  : Integer   ;
begin
  Result := 0 ;
  try
    if Supprim and not ExisteSQL('SELECT * FROM REGION WHERE RG_PAYS="' + RG_PAYS.Text + '" '
                + 'AND RG_REGION<>"' + FListe.Cells[2,Fliste.Row] + '"' )
      then ExecuteSql('UPDATE PAYS SET PY_REGION="-" WHERE PY_PAYS="' + RG_PAYS.Text + '"') ;
  except
    Result := 1 ;
  end ;
end;

procedure TOF_YYREGION.Imprimer(Sender : TObject) ;
begin
  PrintDBGrid (FListe,Nil, 'REGIONS','');
end;

procedure TOF_YYREGION.FListeRowEnter(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean) ;
begin
Modif := False ;
DernLigne := Fliste.row;
BoutonsNavMaj;
end;


procedure TOF_YYREGION.FListeRowExit(Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean) ;
begin

if not EstRemplit(Fliste.row) then Exit
else SauvGrille;

Cancel := not EnregLigne(sender);
end;

function TOF_YYREGION.VerifLigne(Fliste:THGrid):boolean;
begin
   Result := FALSE;
   if (trim(FListe.cells[2,Fliste.row])='') and (trim(FListe.cells[1,Fliste.row])<>'') then PGIBox('Vous devez renseigner un libellé.', 'REGIONS' )
   ELSE
   if (trim(Fliste.cells[1,Fliste.row])='') and (trim(FListe.cells[2,Fliste.row])<>'') then PGIBox('Vous devez renseigner un code.', 'REGIONS')
   else
   if (trim(Fliste.cells[1,Fliste.row])<>'') and (trim(FListe.cells[2,Fliste.row])<>'') then Result := TRUE;
end;


procedure TOF_YYREGION.FListeKeyPress( Sender : TObject ; var Key : Char ) ;
begin
if Key <> #9 then Modif := True ;
end;

{----------------------------------------------------
Suite ........ :   Ctrl-I             :  Impression
Suite ........ :   Ctrl-N ou Inser    :  Insertion
Suite ........ :   Ctrl-suppr         :  Suppression
Suite ........ :   F10                :  validation
Suite ........ :   Entrée dans la Grille : comme Tabulation
Mots clefs ... :
*****************************************************************}
procedure TOF_YYREGION.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var  okG, Vide, AvecCtrl : boolean;
begin
   if RG_PAYS.Text='<<Tous>>' then
      PgiBox('Pour modifier le libellé veuillez sélectionner un pays')
      else
      begin
      if Not FListe.SynEnabled then
         begin
         Key := 0;
         Exit;
         end;
      OkG      := ( Screen.ActiveControl = FListe ) ;
      Vide     := ( Shift = [] ) ;
      AvecCtrl := ( Shift = [ssctrl] ) ;
      // Validation cellule = tabulation
      if (Key = VK_RETURN) then
         begin
         if ((OkG) and (Vide)) then KEY := VK_TAB ;
         end    // Validation ligne
         else
      if Key =VK_F10 then
         begin
         if Vide then
            begin
            Key := 0 ;
            EnregLigne(Fliste);
            end ;
         end
         else
      if Key=VK_DELETE then  // suppresion
         begin
         if ((OkG) and AvecCtrl) then
            begin
            Key := 0 ;
            BDeleteClick(sender);{JP 17/08/04: maintenant il s'agit d'une méthode et non plus d'une évènement}
            end ;
         end
         else
      if (Key=VK_INSERT) or (Key=78) then // nouvelle ligne
         begin
         if (Vide and (Key=VK_INSERT)) or (AvecCtrl and (Key=78)) then
            begin
            Key := 0 ;
            BInsertClick (Sender);
            end ;
         end
         else
         // Impression
      if Key=73 then{^I}
         begin
         if AvecCtrl then
            begin
            Key := 0 ;
            Imprimer(nil) ;
            end ;
         end
         else
         begin
         Action := 'MODIF';  FormKdown:=1;
         end;
   end;
end;


procedure TOF_YYREGION.ParametrageFiche;
begin
  // Param de base
  TitreEcran      := 'Liste des régions' ;
  NomTable        := 'REGION' ;
  ListeChamps     := 'RG_REGION;RG_LIBELLE' ;
  CodeEtat        := '' ;
  NatureEtat      := '' ;
  ListeColsUniq   := '1';
  ChampTri        := 'RG_REGION' ;
end;


procedure TOF_YYREGION.ParametrageGrille(GRILLE: THGrid);
begin
  Grille.ColLengths[1] := 3;         //Code 3 car maxi.
  Grille.ColLengths[2] := 35;        //Libelle sur 35
end;


procedure TOF_YYREGION.LancerRecherche(Pays:string);
var {$IFNDEF EAGLCLIENT}
    Q : TQuery;
    {$ENDIF}
    Sql : string;

begin
  TOBLigne := TOB.Create('Lignes de la liste',nil,-1);
  Sql    := 'SELECT * FROM ' + NomTable ;
  if (Pays <> '') and (Pays<>'<<Tous>>') then
     SQL := SQL + ' WHERE RG_PAYS="' + Pays +'"' ;
  if Trim(ChampTri) <> '' then
     SQL := SQL + ' ORDER BY ' + ChampTri ;
  TOBLigne.ClearDetail;
  // chargement des données
  {$IFDEF EAGLCLIENT}
    MaxRow := (FListe.Height div FListe.DefaultRowHeight) ;
    TOBLigne := OpenSql( SQL, True, MaxRow );
    // MAJ boutons fetch
    SetControlEnabled('BNextQ' , Not TOBLigne.GetValue('EOF') ) ;
    SetControlEnabled('BPrevQ', FListe.TopRow > 1 ) ;
  {$ELSE}
    Q := OpenSql( SQL, True );
    TOBLigne.ClearDetail;
    TOBLigne.LoadDetailDB(NomTable,'','',Q,False);
    Ferme(Q);
  {$ENDIF}

  // Positionnement initiale
  FListe.Col := 2;
  FListe.Row := 1;

  // MAJ Affichage Grille
  FListe.rowCount := 2 ;
  RafraichirGrille(TOBLigne);
//  if TOBLigne.Detail.Count > 0 then {JP 17/08/04: maintenant il s'agit d'une méthode et non plus d'une évènement}
//  TOBLigne.Free;
end;


procedure TOF_YYREGION.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  pgiinfo('GrilleCellEnter');
  pgiinfo('Numéro de ligne'+inttostr(arow));
  DernLigne:= arow;
  if FListe.col = 2  then Exit ;
  if Action='INSERT' then cancel:=FALSE
  else cancel:=TRUE;
end;

procedure TOF_YYREGION.RafraichirGrille(TOBLigne:TOB);
begin
    TOBLigne.PutGridDetail( FListe, False, False, ListeChamps, TRUE ) ;
end;

PROCEDURE TOF_YYREGION.BValider(Sender: TOBject);
begin
    inherited;
   PgiInfo('BValider');
   if VerifLigne(Fliste) then EnregLigne(Sender);

end;

Function TOF_YYREGION.EnregLigne(Sender: TOBject):boolean;
var Rep      : Integer ;
begin
if not EstRemplit( Fliste.row ) then Exit ;
If Modif then
   begin
      Rep:= PGIAskCancel('Voulez-vous enregistrer les modifications ?',Ecran.Caption);
      Case Rep of
           mrYes : try
                      if ExisteEnreg(Fliste,DernLigne) then
                         raise EAbort.Create('')
                      else OnUpdateRecord(sender);
                      if (OnAfterUpdateRecord(sender)<>0) then
                         raise EAbort.Create('') ;
                   except //YCP OK LOGS1
                   on E: Exception do
                     begin
                     if E.Message <> '' then
                        PGIBox( 'Erreur à l''enregistrement.' + #10#13 + E.Message , Ecran.Caption ) ;
                     Result := False ;
                     end;
                   end;
           mrNo  :  begin
                    LancerRecherche(RG_PAYS.Values[RG_PAYS.ItemIndex]);
                    Modif := False ;
                    Result := True;
                    end;
           mrCancel : Result := False ;
           end;
     end;
end;

Function TOF_YYREGION.ExisteEnreg(Fliste:THGrid;Ligne:integer):boolean;
begin
    Result := True;
    if ExisteSql('Select * from Region where RG_PAYS="'+ RG_PAYS.Values[RG_PAYS.ItemIndex]
       + '" and RG_REGION="'+Fliste.cells[1,Ligne]+'" and and RG_LIBELLE="'+Fliste.cells[2,Ligne]+'"') then
       pgibox('L''enregistrement que vous avez saisi existe déjà. Vous devez le modifier.','REGION')
    else  Result := FALSE;
end;

procedure TOF_YYREGION.BoutonsNavMaj;
begin
// MAJ des boutons de navigations
  SetControlEnabled('BFirst',FListe.Row > 1);
  SetControlEnabled('BPrev', FListe.Row > 1);
  SetControlEnabled('BNext', FListe.Row < FListe.RowCount );
  SetControlEnabled('BLast', FListe.Row < FListe.RowCount );
end;

procedure TOF_YYREGION.ParcoursListe(Sender: TObject);
var
  cName : String; // nom du controle
  Row, newRow : Integer;
begin
  cName := TButton(Sender).name;
  Row := FListe.Row;
  newRow := FListe.Row;                 
  if cName = 'BFIRST' then newRow := 1 // Première ligne
  else
  if cName = 'BPREV' then // Ligne précédente
     begin
     if FListe.Row > 1 then newRow := FListe.Row - 1;
     end 
  else
  if cName = 'BNEXT' then // Ligne suivante
     begin	  
    	if (FListe.Row < FListe.RowCount) then newRow := FListe.Row + 1;
     end
  else
  if cName = 'BLAST' then newRow := Fliste.RowCount; // Dernière ligne
  if newRow <> Row then FListe.gotoRow(newRow); // Déplacement
end;

function TOF_YYREGION.Supprim:boolean;
var TobSuppr : TOB ;
    RG_REGION,PAYS,RG_LIBELLE : string;
begin
  if Fliste.RowCount=0 then Exit ;
  RG_REGION := Fliste.cells[1,Fliste.row];
  RG_LIBELLE := Fliste.cells[2,Fliste.row];
  PAYS := RG_PAYS.Values[RG_PAYS.ItemIndex];
  // Message avertissement
  if PGIAskCancel('Voulez-vous supprimez l''enregistrement courant ?', Ecran.Caption) <> mrYes then
     Exit
  else
     try
      ExecuteSql('DELETE from REGION where RG_REGION="'+RG_REGION+'" and RG_LIBELLE="'+RG_LIBELLE+'" and RG_PAYS="'+PAYS+'"');
      PGIInfo('Suppression effectuée',Ecran.caption);
      except
          on E:Exception do
          begin
          if E.Message <> '' then
            PGIBox( 'Erreur à la suppression de l''enregistrement.' + #10#13 + E.Message , Ecran.Caption ) ;
          end ;
      end ;
    LancerRecherche(PAYS);

end;

{$IFDEF EAGLCLIENT}
{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : Gestion du chargement par page des données en eAGL
Mots clefs ... : 
*****************************************************************}
procedure TOF_YYREGION.BNextQClick(Sender: TObject);
var Add,Cpt,Row  : Integer ;
begin

	// Validation des données si nécessaire
  if Sender <> nil then
  if not EnregLigne(FListe.Row,True) then Exit;

  Add := AppendSQL( TOBLignes , MaxRow ) ;
  if Add > 0 then
    begin
    for Cpt := 0 to (Add - 1) do
      begin
      FListe.RowCount := FListe.RowCount + 1 ;
      // indexs enreg dans Q
      Row := TOBLignes.Detail.Count - Add + Cpt ;
      // Affichage Grille
      TOBLignes.Detail[ Row ].PutLigneGrid( FListe , FListe.RowCount - 1 , false , false, ListeChamps ) ;
      // MAJ TOB
      end ;
    end ;
  SetControlEnabled('BNextQ' , Not TOBLignes.GetValue('EOF') ) ;
  SetControlEnabled('BPrevQ', FListe.TopRow > 1 ) ;
end;
{$ENDIF}

{$IFDEF EAGLCLIENT}
procedure TOF_YYREGION.BPrevQClick(Sender: TObject);
begin
  SendNotifyMessage( TWinControl(FListe).Handle, WM_KeyDown, 33, 0 ) ;
end;

{$ENDIF}

Procedure TOF_YYREGION.SauvGrille;
var TobTmp : TOB;
begin
    TOBSauv.ClearDetail ;
    TOBSauv.AddChampSupValeur('PAYS',RG_PAYS.text);
    TOBSauv.GetGridDetail(Fliste,Fliste.Rowcount,NomTable,'RG_PAYS;RG_REGION;RG_LIBELLE',true);
end;

Initialization
  registerclasses ( [ TOF_YYREGION ] ) ;
end.
