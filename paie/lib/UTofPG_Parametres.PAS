{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPARAMETRES ()
Mots clefs ... : TOF;PGPARAMETRES
*****************************************************************}
{
PT1   : 22/02/2007 Intégration de la gestion des "Zones Libres Bulletin"
PT2   : 07/03/2007 FC Rajout du type élément dynamique montant
PT3   : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT4   : 20/03/2007 FC V_72 Rajout du type donnée fiche salarié
PT5   : 22/03/2007 GGU V_72 Gestion de l'onglet Présence
PT6   : 02/04/2007 GGU V_72 Gestion du lookup avec la liste des parametres pour
                            l'onglet de présence
PT7   : 03/04/2007 GGU V_72 Mise à jour des données salariés "Idem population"
                            //Modif le 17/04/07
PT8   : 18/04/2007 GGU V_72 Gestion de la convention dans les données salariés
PT9   : 06/06/2007 FC V_72 Rendre le libellé modifiable en modification
PT10  : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Caractéristique
                           activité" - FQ N°14568
PT11  : 24/08/2007 NA V_8  Filtre la tablette "Module" pour tenir compte de la sérialisation du module "Presence"
}
Unit UTofPG_Parametres ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HCtrls, HEnt1,
{$IFNDEF EAGLCLIENT}
     db, mul, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$else}
     eMul,
{$ENDIF}
     HMsgBox, UTOF, UTOB, utobdebug, HTB97,commun,ENTPAIE, ParamSoc ;

Type
  TOF_PGPARAMETRES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Action, Key : String;
    CEG, STD, DOS : Boolean; //Autorisations d'accès
    TobParametres : TOB;
    //Controles
    BtnIns_l, BtnDel_l, BtnAdd_l, Btn_Delete : TToolbarButton97;
    GD_GRILLEFRAC : THGrid;
    //Controles de l'onglet Caractéristiques
    CB_TYPEPARAMETRE, CB_PREDEFINI, CB_MODULECHAMP : THValComboBox;
    Ed_PGPARAMETRE, Ed_LIBELLE : THEdit;
    //Onglets paramètres
    Tab_PROFIL, Tab_CP, Tab_BANQUE, Tab_PRESENCE, Tab_FRACTION, Tab_ZONES1, Tab_ZONES2 : TTabSheet;
    //Valeurs d'initialisation
    init_PGPARAMETRE, init_TYPEPARAMETRE, init_PREDEFINI, init_LIBELLE, init_MODULECHAMP : String;  // , init_NODOSSIER
    procedure OnMODULECHAMPExit(Sender : TObject);
    procedure OnMODULECHAMPEnter(Sender : TObject);
    procedure OnPREDEFINIExit(Sender : TObject);
    Procedure UpdateVisibilites;
    Function GetNoDossier(Predef : String) : String;
    Procedure OnDeleteClick(Sender : TObject);
    procedure TypeZoneChange(Sender: TObject);
    procedure CheckValuesBUL(Sender: TObject);
    procedure BtnAddClick(Sender: TObject);
    procedure BtnDelClick(Sender: TObject);
    procedure BtnInsClick(Sender: TObject);
    procedure ShowLookupPROFILPRES(Sender : TObject);
    procedure OnExitActivite(Sender: TObject);
    procedure ChargeProfilActivite(LeChamp, Nat: string);
    procedure ChargementTabZonesLibres();
    procedure ChargementTabZonesSalaries();
  end ;

Implementation

Uses
  StrUtils, PGOutils, PGOutils2, Grids, lookup, ed_Tools, pgpopuloutils {PT7};

procedure TOF_PGPARAMETRES.OnNew ;
begin
  Inherited ;
//  SetControlText('PGP_MODULECHAMP','PAI');
end ;

procedure TOF_PGPARAMETRES.OnDelete ;
begin
  Inherited ;
  ExecuteSQL('DELETE FROM PGPARAMETRES WHERE ##PGP_PREDEFINI## AND '+Key);
end ;

procedure TOF_PGPARAMETRES.OnUpdate;
  Procedure DeleteGridParametre(ColName : String);
  begin
    ExecuteSQL('DELETE FROM PGPARAMETRES WHERE ##PGP_PREDEFINI## AND PGP_PGNOMCHAMP LIKE "'+ColName+'%" AND '+Key);
  end;

  procedure InsertOrUpdateParametre(Champs, Valeur : String);
  var
    Value : String;
    NewTob : TOB;
  begin
    //On recherche si l'enregistrement existe déjà,
    NewTob := TobParametres.FindFirst(['PGP_PGPARAMETRE','PGP_TYPEPARAMETRE','PGP_PREDEFINI','PGP_NODOSSIER','PGP_MODULECHAMP','PGP_PGNOMCHAMP']
                                      ,[init_PGPARAMETRE,init_TYPEPARAMETRE,init_PREDEFINI,GetNoDossier(init_PREDEFINI),init_MODULECHAMP,Champs]
                                      ,True);
    if Assigned(NewTob) then
    begin
    //Si oui, on le modifie
      NewTob.PutValue('PGP_PGVALCHAMP',Valeur);
    end else begin
    //Sinon, on le crée
      Value := GetControlText('PGP_PGPARAMETRE')   +'##S&p@r@t3vr##'
             + GetControlText('PGP_TYPEPARAMETRE') +'##S&p@r@t3vr##'
             + GetNoDossier(GetControlText('PGP_PREDEFINI'))+'##S&p@r@t3vr##'
             + GetControlText('PGP_PREDEFINI')     +'##S&p@r@t3vr##'
             + GetControlText('PGP_LIBELLE')       +'##S&p@r@t3vr##'
             + GetControlText('PGP_MODULECHAMP')   +'##S&p@r@t3vr##'
             + Champs                              +'##S&p@r@t3vr##'
             + Valeur;
      NewTob := TOB.Create('PGPARAMETRES',TobParametres,-1);
      NewTob.LoadFromSt('PGP_PGPARAMETRE,PGP_TYPEPARAMETRE,PGP_NODOSSIER,PGP_PREDEFINI,PGP_LIBELLE,PGP_MODULECHAMP,PGP_PGNOMCHAMP,PGP_PGVALCHAMP',
                                               ',', Value, '##S&p@r@t3vr##');
    end;
  end;

  Procedure SaveControls(Conteneur : TWinControl);
  var
    i, indexCol, indexRow : integer;
    TempGrid : THGrid;
  begin
    InitMoveProgressForm(nil,Ecran.Caption,TraduireMemoire('Sauvegarde en cours, veuillez patienter...'),Conteneur.ControlCount,False,True);
    for i := 0 to Conteneur.ControlCount -1 do
    begin
      MoveCurProgressForm(TraduireMemoire('Conteneur ')+Conteneur.Name+' - '+IntToStr(i));
      //Recursion pour les objets conteneurs
      if (Conteneur.Controls[i] is TWinControl) and ((Conteneur.Controls[i] as TWinControl).ControlCount > 0) then
        SaveControls(Conteneur.Controls[i] as TWinControl);
      if      (Conteneur.Controls[i] is THEdit)             then
        InsertOrUpdateParametre(Conteneur.Controls[i].Name,(Conteneur.Controls[i] as THEdit).Text)
      else if (Conteneur.Controls[i] is THValComboBox)      then
        InsertOrUpdateParametre(Conteneur.Controls[i].Name,(Conteneur.Controls[i] as THValComboBox).Value)
      else if (Conteneur.Controls[i] is THNumEdit)          then
        InsertOrUpdateParametre(Conteneur.Controls[i].Name,(Conteneur.Controls[i] as THNumEdit).Text)
      else if (Conteneur.Controls[i] is THSpinEdit)         then
        InsertOrUpdateParametre(Conteneur.Controls[i].Name,(Conteneur.Controls[i] as THSpinEdit).Text)
      else if (Conteneur.Controls[i] is THMultiValComboBox) then
        InsertOrUpdateParametre(Conteneur.Controls[i].Name,(Conteneur.Controls[i] as THMultiValComboBox).Value)
      else if (Conteneur.Controls[i] is THGrid)             then
      begin
        TempGrid := (Conteneur.Controls[i] as THGrid);
        for indexCol := TempGrid.FixedCols to TempGrid.ColCount-1 do
        begin
          { Suppression des lignes pré-existante }
          DeleteGridParametre(TempGrid.ColNames[indexCol]+';');
          { Sauvegarde des nouvelles lignes de la grid }
          for indexRow := TempGrid.FixedRows to TempGrid.RowCount-1 do
            InsertOrUpdateParametre(TempGrid.ColNames[indexCol]+';'+intToStr(indexRow),TempGrid.CellValues[indexCol,indexRow]);
        end;
      end;
    end;
    FiniMoveProgressForm;
  end;

var
  i : integer;
  Qry : TQuery;
  ObjPopulations : TUpdateIdemPop;
  St : String;
  TmpStringList : TStringList;
begin
  Inherited ;
  LastErrorMsg := '';

  if (CB_TYPEPARAMETRE.Value = '') then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; CB_TYPEPARAMETRE.SetFocus;
    LastErrorMsg := TraduireMemoire('Vous devez choisir le type de paramètres.');
    LastError := 1;
    exit;
  end;
  if (CB_PREDEFINI.Value = '') then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; CB_PREDEFINI.SetFocus;
    LastErrorMsg := TraduireMemoire('Vous devez choisir le type de prédéfini.');
    LastError := 2;
    exit;
  end else if (CB_PREDEFINI.Value = 'CEG') and (CEG = FALSE) then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; CB_PREDEFINI.SetFocus;
    SetControlText('PGP_PREDEFINI','DOS');
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas créer de prédéfini CEGID.');
    LastError := 3;
    exit;
  end else if (CB_PREDEFINI.Value = 'STD') and (STD = FALSE) then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; CB_PREDEFINI.SetFocus;
    SetControlText('PGP_PREDEFINI','DOS');
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas créer de prédéfini Standard.');
    LastError := 4;
    exit;
  end;
  if (CB_MODULECHAMP.Value = '') then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; CB_MODULECHAMP.SetFocus;
    LastErrorMsg := TraduireMemoire('Vous devez choisir le module associé.');
    LastError := 5;
    exit;
  end;
  if (CompareStr(LeftStr(Ed_PGPARAMETRE.Text,3),CB_MODULECHAMP.Value) <> 0) then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; Ed_PGPARAMETRE.SetFocus;
    Ed_PGPARAMETRE.Text := CB_MODULECHAMP.Value+RightStr(Ed_PGPARAMETRE.Text,Length(Ed_PGPARAMETRE.Text)-3);
  end;
  if (length(Ed_PGPARAMETRE.Text) <> 6) then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; Ed_PGPARAMETRE.SetFocus;
    LastErrorMsg := TraduireMemoire('Vous devez renseigner un code paramètre composé de 3 lettres (dépendant du module associé) et de 3 chiffres.');
    LastError := 6;
    exit;
  end;
  if (CB_PREDEFINI.Value = 'CEG') and (not (StrToInt(RightStr(Ed_PGPARAMETRE.Text,1)) mod 2 = 0) ) then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; Ed_PGPARAMETRE.SetFocus;
    LastErrorMsg := TraduireMemoire('Le code paramètre doit finir par un chiffre PAIR.');
    LastError := 7;
    Exit;
  end;
  if (CB_PREDEFINI.Value = 'STD') and
     (    (    (StrToInt(RightStr(Ed_PGPARAMETRE.Text,1)) <> 1 )
           and (StrToInt(RightStr(Ed_PGPARAMETRE.Text,1)) <> 3 )
          )
     ) then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; Ed_PGPARAMETRE.SetFocus;
    LastErrorMsg := TraduireMemoire('Le code paramètre doit finir par 1 ou 3.');
    LastError := 8;
    Exit;
  end;
  if (CB_PREDEFINI.Value = 'DOS') and
     ( (    (StrToInt(RightStr(Ed_PGPARAMETRE.Text,1)) <> 5)
        and (StrToInt(RightStr(Ed_PGPARAMETRE.Text,1)) <> 7)
        and (StrToInt(RightStr(Ed_PGPARAMETRE.Text,1)) <> 9)
       )
     ) then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; Ed_PGPARAMETRE.SetFocus;
    LastErrorMsg := TraduireMemoire('Le code paramètre doit finir par 5, 7 ou 9.');
    LastError := 9;
    Exit;
  end;
  if    (Action = 'CREATION')
     and ExisteSQL('select pgp_pgparametre from pgparametres where pgp_pgparametre = "'+Ed_PGPARAMETRE.Text+'"') then
  begin
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 0; Ed_PGPARAMETRE.SetFocus;
    LastErrorMsg := TraduireMemoire('Ce code paramètre existe déjà.');
    LastError := 10;
    Exit;
  end;

  if LastErrorMsg = '' then
  begin
    for i := 1 to 24 do
    begin
      if (GetControl('OPERATEUR' + IntToStr(i)) as THValComboBox).Visible then
      begin
        CheckValuesBUL((GetControl('OPERATEUR' + IntToStr(i)) as THValComboBox));
        if LastErrorMsg <> '' then Break;
      end;
      if (GetControl('CBVALZONE' + IntToStr(i)) as THValComboBox).Visible then
      begin
        CheckValuesBUL((GetControl('CBVALZONE' + IntToStr(i)) as THValComboBox));
        if LastErrorMsg <> '' then Break;
      end;
      if (GetControl('CBVALZONE' + IntToStr(i) + 'B') as THValComboBox).Visible then
      begin
        CheckValuesBUL((GetControl('CBVALZONE' + IntToStr(i) + 'B') as THValComboBox));
        if LastErrorMsg <> '' then Break;
      end;
      if (GetControl('VALZONE' + IntToStr(i)) as THEdit).Visible then
      begin
        CheckValuesBUL((GetControl('VALZONE' + IntToStr(i)) as THEdit));
        if LastErrorMsg <> '' then Break;
      end;
    end;
  end;

  if LastErrorMsg = '' then
  begin
    if      (CB_MODULECHAMP.Value = 'CPA') then
      SaveControls(Tab_CP)
    else if (CB_MODULECHAMP.Value = 'ACF') then
      SaveControls(Tab_FRACTION)
    else if (CB_MODULECHAMP.Value = 'PRE') then   //PT5
      SaveControls(Tab_PRESENCE)
    else if (CB_MODULECHAMP.Value = 'SAL') then
      SaveControls(Tab_PROFIL)
    else if (CB_MODULECHAMP.Value = 'BQE') then
      SaveControls(Tab_BANQUE)
    else if (CB_MODULECHAMP.Value = 'PAI') then
    begin
      SaveControls(Tab_ZONES1);
      SaveControls(Tab_ZONES2);
    end;
    TobParametres.SetAllModifie(True);
    TobParametres.InsertOrUpdateDB;
    
    //Debut PT7
    //On recherche à quelles populations est associé le paramètre.
{$IFDEF IDEMPOP}
    Qry:=OpenSQL('SELECT DISTINCT PGO_CODEASSOCIE FROM PGPARAMETRESASSOC '
                +'where ##PGO_PREDEFINI## and PGO_TYPEPARAMETRE = "POP" '
                +' and PGO_PGPARAMETRE = "'+Ed_PGPARAMETRE.Text+'" ',True) ;
    ObjPopulations := TUpdateIdemPop.Create;
    While Not Qry.EOF do
      BEGIN
      //Pour chaque population associée, on met a jour les valeurs "Idem population" des
      //salariés concernés
      TmpStringList := ObjPopulations.MajDonneesSalariePop(Qry.Fields[0].AsString,Date());
      FreeAndNil(TmpStringList);
      Qry.Next ;
      END ;
    Ferme(Qry) ;
    ObjPopulations.Free;
{$ENDIF}
    //Fin PT7

    //DEB PT9
    st := 'UPDATE PGPARAMETRES SET PGP_LIBELLE="' + GetControlText('PGP_LIBELLE') + '"'
      + ' WHERE ' + Key;
    ExecuteSql(St);
    //DEB PT9

    Ecran.Close;
  end else begin
    //Les erreurs ont déjà été signalées donc on réinitialise les variables d'erreur
    LastErrorMsg := '';
    LastError := 0;
  end;
end ;

procedure TOF_PGPARAMETRES.OnLoad;
//var
// MaxCol, MaxRow : Integer;
  Function  FindValue(Champs : String) : String;
  var
    NewTob : TOB;
  begin
    NewTob := TobParametres.FindFirst( ['PGP_PGPARAMETRE','PGP_TYPEPARAMETRE','PGP_PREDEFINI','PGP_NODOSSIER','PGP_MODULECHAMP','PGP_PGNOMCHAMP']
                                      ,[init_PGPARAMETRE,init_TYPEPARAMETRE,init_PREDEFINI,GetNoDossier(init_PREDEFINI),init_MODULECHAMP,Champs]
                                      ,True);
    if Assigned(NewTob) then
      result := NewTob.GetValue('PGP_PGVALCHAMP')
    else
      result := '';
  end;

  Function  FindMaxValues(Champs : String; initial : Integer) : Integer;
  var
    NewTob : TOB;
    indexRow : Integer;
  begin
    indexRow := initial;
    Result := initial;
    NewTob := TobParametres.FindFirst( ['PGP_PGPARAMETRE','PGP_TYPEPARAMETRE','PGP_PREDEFINI','PGP_NODOSSIER','PGP_MODULECHAMP','PGP_PGNOMCHAMP']
                                      ,[init_PGPARAMETRE,init_TYPEPARAMETRE,init_PREDEFINI,GetNoDossier(init_PREDEFINI),init_MODULECHAMP,Champs+';'+intToStr(indexRow)]
                                      ,True);
    while Assigned(NewTob) do
    begin
      Result := indexRow;
      Inc(indexRow);
      NewTob := TobParametres.FindFirst(['PGP_PGPARAMETRE','PGP_TYPEPARAMETRE','PGP_PREDEFINI','PGP_NODOSSIER','PGP_MODULECHAMP','PGP_PGNOMCHAMP']
                                       ,[init_PGPARAMETRE,init_TYPEPARAMETRE,init_PREDEFINI,GetNoDossier(init_PREDEFINI),init_MODULECHAMP,Champs+';'+intToStr(indexRow)]
                                       ,True);
    end;
  end;

  Procedure LoadControls(Conteneur : TWinControl);
   var
     i, indexCol, indexRow : integer;//, MaxRow
     TempGrid : THGrid;
     TempStr : String;
     TempTHEdit : THEdit;
     TempTHNumEdit : THNumEdit;
     TempTHSpinEdit : THSpinEdit;
     TempTHValComboBox : THValComboBox;
     TempTHMultiValComboBox : THMultiValComboBox;
   begin
     for i := 0 to Conteneur.ControlCount -1 do
     begin
       //Recursion pour les objets conteneurs
       if (Conteneur.Controls[i] is TWinControl) and ((Conteneur.Controls[i] as TWinControl).ControlCount > 0) then
       begin
         LoadControls(Conteneur.Controls[i] as TWinControl);
       end;
       if      (Conteneur.Controls[i] is THEdit) then
       begin
         TempTHEdit := (Conteneur.Controls[i] as THEdit);
         TempStr := FindValue(TempTHEdit.Name);
         if TempStr <> '' then TempTHEdit.Visible := True;
         TempTHEdit.Text := TempStr;
       end else if (Conteneur.Controls[i] is THNumEdit) then
       begin
         TempTHNumEdit := (Conteneur.Controls[i] as THNumEdit);
         TempStr := FindValue(TempTHNumEdit.Name);
         if TempStr <> '' then TempTHNumEdit.Visible := True;
         TempTHNumEdit.Text := TempStr;
       end else if (Conteneur.Controls[i] is THSpinEdit) then
       begin
         TempTHSpinEdit := (Conteneur.Controls[i] as THSpinEdit);
         TempStr := FindValue(TempTHSpinEdit.Name);
         if TempStr <> '' then TempTHSpinEdit.Visible := True;
         TempTHSpinEdit.Text := TempStr;
       end else if (Conteneur.Controls[i] is THValComboBox) then
       begin
         TempTHValComboBox := (Conteneur.Controls[i] as THValComboBox);
         TempStr := FindValue(TempTHValComboBox.Name);
         if TempStr <> '' then TempTHValComboBox.Visible := True;
         TempTHValComboBox.Value := TempStr;
       end else if (Conteneur.Controls[i] is THMultiValComboBox) then
       begin
         TempTHMultiValComboBox := (Conteneur.Controls[i] as THMultiValComboBox);
         TempStr := FindValue(TempTHMultiValComboBox.Name);
         if TempStr <> '' then TempTHMultiValComboBox.Visible := True;
         TempTHMultiValComboBox.Value := TempStr;
       end else if (Conteneur.Controls[i] is THGrid)             then
       begin
         TempGrid := (Conteneur.Controls[i] as THGrid);
         for indexCol := 0 to TempGrid.ColCount-1 do
         begin
           TempGrid.RowCount := FindMaxValues(TempGrid.ColNames[indexCol],TempGrid.FixedRows)+1;
           for indexRow := TempGrid.FixedRows to TempGrid.RowCount-1 do
           begin
             TempGrid.CellValues[indexCol, indexRow] := FindValue(TempGrid.ColNames[indexCol]+';'+intToStr(indexRow))
           end;
         end;
       end
     end;
   end;

begin
  Inherited ;

  Key := '(     PGP_PGPARAMETRE   ="'+Init_PGPARAMETRE   +'"'
        +'  AND PGP_NODOSSIER     ="'+GetNoDossier(init_PREDEFINI) +'"'
        +'  AND PGP_TYPEPARAMETRE ="'+init_TYPEPARAMETRE +'"'
        +'  AND PGP_PREDEFINI     ="'+init_PREDEFINI     +'"'
        +'  AND PGP_MODULECHAMP   ="'+init_MODULECHAMP   +'")';
  if Assigned(TobParametres) then FreeAndNil(TobParametres);
  TobParametres := TOB.Create('Paramètres', nil, -1);
  TobParametres.ClearDetail;
  TobParametres.LoadDetailDBFromSQL('PGPARAMETRES','SELECT * FROM PGPARAMETRES WHERE ##PGP_PREDEFINI## '+Key);
  if      (CB_MODULECHAMP.Value = 'CPA') then
    LoadControls(Tab_CP)
  else if (CB_MODULECHAMP.Value = 'ACF') then
    LoadControls(Tab_FRACTION)
  else if (CB_MODULECHAMP.Value = 'PRE') then //PT5
    LoadControls(Tab_PRESENCE)
  else if (CB_MODULECHAMP.Value = 'SAL') then
    LoadControls(Tab_PROFIL)
  else if (CB_MODULECHAMP.Value = 'BQE') then
    LoadControls(Tab_BANQUE)
  else if (CB_MODULECHAMP.Value = 'PAI') then
  begin
    LoadControls(Tab_ZONES1);
    LoadControls(Tab_ZONES2);
  end;
  UpdateVisibilites;
  OnExitActivite(Self);//PT8
end ;

procedure TOF_PGPARAMETRES.OnArgument (S : String ) ;
var
  Arg, Stplus : String;      // pt11
  i :Integer;
  TempBo : Boolean;
begin
  Inherited ;
  Arg := S;
  Action := Trim(ReadTokenPipe(Arg, ';'));
  if (Action = 'ACTION=MODIFICATION') then Action := 'MODIFICATION' else Action := 'CREATION';

  // deb pt11
  stplus := '';
  if not VH_PAIE.PGSeriaPresence then  stplus := stplus + ' AND CO_CODE <> "PRE"';
  if stplus <> '' then SetControlProperty('PGP_MODULECHAMP', 'Plus', stplus);
  // fin pt11
  
  init_PGPARAMETRE   := Trim(ReadTokenPipe(Arg, ';'));
  init_PREDEFINI     := Trim(ReadTokenPipe(Arg, ';'));
  init_TYPEPARAMETRE := Trim(ReadTokenPipe(Arg, ';'));
  init_LIBELLE       := Trim(ReadTokenPipe(Arg, ';'));
  init_MODULECHAMP   := Trim(ReadTokenPipe(Arg, ';'));
//  if init_MODULECHAMP = '' then init_MODULECHAMP := 'PAI';
  if init_TYPEPARAMETRE = '' then init_TYPEPARAMETRE := 'POP';
  if init_PREDEFINI = '' then init_PREDEFINI := 'STD';
  AccesPredefini('TOUS', CEG, STD, DOS);
  Tab_PROFIL       := (GetControl('TBPROFIL')          as TTabSheet);
  Tab_CP           := (GetControl('TBCP')              as TTabSheet);
  Tab_PRESENCE     := (GetControl('TBPRESENCE')        as TTabSheet);  //PT5
  Tab_BANQUE       := (GetControl('TBBANQUE')          as TTabSheet);
  Tab_FRACTION     := (GetControl('GBFRACTIONNEMENT')  as TTabSheet);
  Tab_ZONES1       := (GetControl('TBZONES1')          as TTabSheet);
  Tab_ZONES2       := (GetControl('TBZONES2')          as TTabSheet);
  CB_TYPEPARAMETRE := (GetControl('PGP_TYPEPARAMETRE') as THValComboBox);
  CB_PREDEFINI     := (GetControl('PGP_PREDEFINI')     as THValComboBox);
  CB_MODULECHAMP   := (GetControl('PGP_MODULECHAMP')   as THValComboBox);
  Ed_PGPARAMETRE   := (GetControl('PGP_PGPARAMETRE')   as THEdit);
  Ed_LIBELLE       := (GetControl('PGP_LIBELLE')       as THEdit);
  Btn_Delete       := (GetControl('BDelete')           as TToolbarButton97);
  BtnIns_l := (GetControl('BINSERT_LINE') as TToolBarButton97);
  BtnDel_l := (GetControl('BDEL_LINE')    as TToolBarButton97);
  BtnAdd_l := (GetControl('BADD_LINE')    as TToolBarButton97);
  //Initialisation des champs
  CB_TYPEPARAMETRE.Value := init_TYPEPARAMETRE;
  CB_PREDEFINI.Value := init_PREDEFINI;
  SetControlText('PGP_PGPARAMETRE',init_PGPARAMETRE);
  SetControlText('PGP_LIBELLE',init_LIBELLE);
  SetControlText('PGP_MODULECHAMP',init_MODULECHAMP);
  CB_TYPEPARAMETRE.Enabled := (Action = 'CREATION');
  CB_PREDEFINI.Enabled := (Action = 'CREATION');
  SetControlEnabled('PGP_PGPARAMETRE',(Action = 'CREATION'));
// PT9 SetControlEnabled('PGP_LIBELLE',(Action = 'CREATION'));
  SetControlEnabled('PGP_MODULECHAMP',(Action = 'CREATION'));
  //Affectation d'un nom aux colonnes des Grid, pour stocker les valeurs avec un
  //nom de champs reprenant le nom de la colonne
  GD_GRILLEFRAC := (GetControl('GRILLEFRAC') as THGrid);
  GD_GRILLEFRAC.ColNames[0] := 'GridJoursRestant';
  GD_GRILLEFRAC.ColNames[1] := 'GridJoursAcquis';
  GD_GRILLEFRAC.ColTypes[0] := 'R';
  GD_GRILLEFRAC.ColFormats[0] := '# ##0.00';
  GD_GRILLEFRAC.ColTypes[1] := 'R';
  GD_GRILLEFRAC.ColFormats[1] := '# ##0.00';
  for i := 1 to 24 do
  begin
    //Initialisation des visibilités de l'onglet Zones libres Bulletin
    (GetControl('OPERATEUR' + IntToStr(i)) as THValComboBox).Visible := False;
    (GetControl('CBVALZONE' + IntToStr(i)) as THValComboBox).Visible := False;
    (GetControl('CBVALZONE' + IntToStr(i) + 'B') as THValComboBox).Visible := False;
    (GetControl('VALZONE'   + IntToStr(i)) as THEdit).Visible := False;
    (GetControl('ZLS'       + IntToStr(i)) as THValComboBox).Visible := False;
    (GetControl('ZSAL'      + IntToStr(i)) as THValComboBox).Visible := False;
    //Affectation des évènements
    (GetControl('TYPZONE'   + IntToStr(i)) as THValComboBox).OnChange := TypeZoneChange;
  end;
  //Activation ou non des boutons et des onglets en fonction des Prédéfinis et
  //des droits de l'utilisateur
  TempBo := (   (CB_PREDEFINI.Value = 'CEG') and (CEG = True)
             or (CB_PREDEFINI.Value = 'STD') and (STD = True)
             or (CB_PREDEFINI.Value = 'DOS') and (DOS = True));
  SetControlEnabled('BValider',TempBo);
  SetControlEnabled('BDelete',TempBo);
  SetControlEnabled('TBPROFIL',TempBo);
  SetControlEnabled('TBPRESENCE',TempBo); //PT5
  SetControlEnabled('TBCP',TempBo);
  SetControlEnabled('TBBANQUE',TempBo);
  SetControlEnabled('GBFRACTIONNEMENT',TempBo);
  SetControlEnabled('TBZONES1',TempBo);
  SetControlEnabled('TBZONES2',TempBo);
  Ed_PGPARAMETRE.OnExit := OnMODULECHAMPExit;
  Ed_PGPARAMETRE.OnEnter := OnMODULECHAMPEnter;
  CB_MODULECHAMP.OnExit := OnMODULECHAMPExit;
  CB_PREDEFINI.OnExit   := OnPREDEFINIExit;
  Btn_Delete.OnClick := OnDeleteClick;
  BtnIns_l.OnClick := BtnInsClick;
  BtnDel_l.OnClick := BtnDelClick;
  BtnAdd_l.OnClick := BtnAddClick;

//PT3
  SetPlusBanqueCP (GetControl ('RIBVIRSOC'));
  SetPlusBanqueCP (GetControl ('RIBACPSOC'));
  SetPlusBanqueCP (GetControl ('RIBFRAISSOC'));
//FIN PT3

  if (CB_MODULECHAMP.Value = 'PRE') then //PT6
    (GetControl('PROFILPRES') as THEdit).OnElipsisClick := ShowLookupPROFILPRES;

  (GetControl('ACTIVITE') as THValComboBox).OnExit := OnExitActivite;  //PT8

  ChargementTabZonesLibres;
  ChargementTabZonesSalaries;
end ;

//Debut PT8
Procedure TOF_PGPARAMETRES.OnExitActivite (Sender: TObject);
begin
  if GetControlText('ACTIVITE') <> '' then
  begin
    ChargeProfilActivite('PROFILAFP', 'AFP');
    ChargeProfilActivite('PROFILANCIEN', 'ANC');
    ChargeProfilActivite('PROFILAPP', 'APP');
    ChargeProfilActivite('PROFILCDD', 'CDD');
    ChargeProfilActivite('PROFILCGE', 'CGE');
    ChargeProfilActivite('PROFILFNAL', 'FNL');
    ChargeProfilActivite('PROFILMUT', 'MUT');
    ChargeProfilActivite('PERIODBUL', 'PER');
    ChargeProfilActivite('PROFILPRE', 'PRE');
    ChargeProfilActivite('PROFIL', 'PRO');
    ChargeProfilActivite('PROFILRBS', 'RBS');
    ChargeProfilActivite('PROFILREM', 'REM');
    ChargeProfilActivite('PROFILRET', 'RET');
    ChargeProfilActivite('REDREPAS', 'RRE');
    ChargeProfilActivite('REDRTT1', 'RT1');
    ChargeProfilActivite('REDRTT2', 'RT2');
    ChargeProfilActivite('PROFILTPS', 'TPS');
    ChargeProfilActivite('PROFILTRANS', 'TRA');
    ChargeProfilActivite('PROFILTSS', 'TSS');
  end;
end;

procedure TOF_PGPARAMETRES.ChargeProfilActivite(LeChamp, Nat: string);
begin
    if GetControlText('ACTIVITE') <> '' then
      SetControlProperty(LeChamp, 'Plus', 'AND (PPI_ACTIVITE="' + GetControlText('ACTIVITE') + '" OR PPI_ACTIVITE="" OR PPI_ACTIVITE IS NULL)')
    else
      SetControlProperty(LeChamp, 'Plus', '');
end;
//Fin PT8

procedure TOF_PGPARAMETRES.OnClose ;
begin
  Inherited ;
  if Assigned(TobParametres) then FreeAndNil(TobParametres);
end ;

procedure TOF_PGPARAMETRES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGPARAMETRES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGPARAMETRES.OnMODULECHAMPExit(Sender: TObject);
begin
  if (CompareStr(LeftStr(Ed_PGPARAMETRE.Text,3),CB_MODULECHAMP.Value) <> 0) then
  begin
    Ed_PGPARAMETRE.Text := CB_MODULECHAMP.Value+RightStr(Ed_PGPARAMETRE.Text,Length(Ed_PGPARAMETRE.Text)-3);
    Ed_PGPARAMETRE.SetFocus;
  end;
  UpdateVisibilites;
end;

procedure TOF_PGPARAMETRES.OnMODULECHAMPEnter(Sender: TObject);
begin
  if (CompareStr(LeftStr(Ed_PGPARAMETRE.Text,3),CB_MODULECHAMP.Value) <> 0) and (CB_MODULECHAMP.Value <> '') then
  begin
    Ed_PGPARAMETRE.Text := CB_MODULECHAMP.Value+RightStr(Ed_PGPARAMETRE.Text,Length(Ed_PGPARAMETRE.Text)-3);
//    Ed_PGPARAMETRE.SetFocus;
  end;
//  UpdateVisibilites;
end;

procedure TOF_PGPARAMETRES.OnPREDEFINIExit(Sender: TObject);
begin
  if (GetControlText('PTE_PREDEFINI') = 'CEG') and (CEG = FALSE) then
  begin
    PGIBox('Vous ne pouvez pas créer de prédéfini CEGID.', Ecran.Caption);
    SetControlText('PTE_PREDEFINI','DOS');
  end;
  if (GetControlText('PTE_PREDEFINI') = 'STD') and (STD = FALSE) then
  begin
    PGIBox('Vous ne pouvez pas créer de prédéfini Standard.', Ecran.Caption);
    SetControlText('PTE_PREDEFINI','DOS');
  end;
end;

procedure TOF_PGPARAMETRES.UpdateVisibilites;
begin
  Tab_CP.TabVisible       := (CB_MODULECHAMP.Value = 'CPA');
  Tab_FRACTION.TabVisible := (CB_MODULECHAMP.Value = 'ACF');
  Tab_PROFIL.TabVisible   := (CB_MODULECHAMP.Value = 'SAL');
  Tab_BANQUE.TabVisible   := (CB_MODULECHAMP.Value = 'BQE');
  Tab_PRESENCE.TabVisible := (CB_MODULECHAMP.Value = 'PRE'); //PT5
  Tab_ZONES1.TabVisible   := (CB_MODULECHAMP.Value = 'PAI');
  Tab_ZONES2.TabVisible   := (CB_MODULECHAMP.Value = 'PAI');
end;

function TOF_PGPARAMETRES.GetNoDossier(Predef : String): String;
begin
  if Predef <> 'DOS' then
   result := '000000'
  else result := PgRendNoDossier;
end;

procedure TOF_PGPARAMETRES.OnDeleteClick(Sender: TObject);
begin
  if PGIAsk(TraduireMemoire('Etes vous sûr de vouloir supprimer ces paramètres de population ?'),Ecran.Caption) <> mrYes then exit;
  OnDelete;
  Ecran.Close;
end;

procedure TOF_PGPARAMETRES.TypeZoneChange(Sender: TObject);
var
  Combo, CbValeur, CbValeurB, Operateur, LValeur, LValeur2 :THValComboBox;
  Valeur :THEdit;
  Libelle:THLabel;
  Indice : String;
begin
  if Sender is TControl then
    Indice := RightStr((Sender as TControl).Name,Length((Sender as TControl).Name)-7);
  Combo     := THValComboBox(GetControl('TYPZONE'   + Indice));
  Valeur    := THEdit       (GetControl('VALZONE'   + Indice));
  CbValeur  := THValComboBox(GetControl('CBVALZONE' + Indice));
  CbValeurB := THValComboBox(GetControl('CBVALZONE' + Indice + 'B'));
  Operateur := THValComboBox(GetControl('OPERATEUR' + Indice));
  Libelle   := THLabel      (GetControl('LBLVALEUR' + Indice));
  LValeur   := THValComboBox(GetControl('ZLS' + Indice));
  LValeur2  := THValComboBox(GetControl('ZSAL' + Indice));
  if (Combo <> nil) and (Combo.Value <> '') and (Valeur <> nil) then
  begin
    Valeur.DataType := '';
    if (Combo.value = 'CMP') or (Combo.value = 'CMM') then
      Valeur.DataType := 'PGCUMULPAIE';
    if (Combo.value = 'BAS') or (Combo.value = 'TAU') or (Combo.value = 'COE') or (Combo.value = 'MNT') then
      Valeur.DataType := 'PGREMUNERATION';
    if (Combo.value = 'ELD') then
      Valeur.DataType := 'PGZONEELTDYN';
    if (Combo.value = 'ELM') then     //PT2
      Valeur.DataType := 'PGELEMENTNAT';
//    if (Combo.Value = 'SAL') then     //PT4
//      Valeur.DataType := 'PGZONESALARIE';

    Valeur.Visible    := (Combo.value <> 'RES');
    Operateur.Visible := (Combo.value = 'RES');
    CbValeur.Visible  := (Combo.value = 'RES');
    CbValeurB.Visible := (Combo.value = 'RES');
    Libelle.Visible   := ((Combo.value <> 'ETQ') and (Combo.value <> 'RES'));
    LValeur.Visible   := (Combo.value = 'ZLS');
    LValeur2.Visible  := (Combo.value = 'SAL');
    if Combo.value <> 'ZLS' then
      LValeur.ItemIndex := -1;
    if Combo.value <> 'SAL' then
      LValeur2.ItemIndex := -1;
    if not Operateur.Visible then Operateur.Value := '';
    if not Valeur.Visible then Valeur.Text     := '';
    if not CbValeur.Visible then CbValeur.Value  := '';
    if not CbValeurB.Visible then CbValeurB.Value := '';

    if (Combo.value = 'ETQ') then
    begin
      Valeur.ElipsisButton := False;
      Valeur.Width := 201;
      Valeur.MaxLength := 17;  //PT2
    end else if (Combo.value = 'ELD') then //DEB PT4  (Combo.value = 'SAL') or  or (Combo.value = 'ZLS')
    begin
      Valeur.ElipsisButton := True;
      Valeur.Width := 120;
      Valeur.MaxLength := 17;
      Libelle.Width := 100;
      Libelle.Left := Valeur.Left + Valeur.Width + 15;   //FIN PT4
    end else if (Combo.value <> 'RES') then
    begin
      Valeur.ElipsisButton := True;
      Valeur.Width := 69;
      Valeur.MaxLength := 0;
      Libelle.Width := 122;   //PT4
      Libelle.Left := Valeur.Left + Valeur.Width + 15;   //PT4
    end;
  end;
end;

procedure TOF_PGPARAMETRES.ChargementTabZonesLibres();
var
  I,J : integer;
  LValeur : THValComboBox;
begin
  for I := 1 to 24 do
  begin
    LValeur := THValComboBox(GetControl('ZLS' + IntToStr(I)));
    if LValeur <> nil then
    begin
      // Zones combos libres
      for J := 1 to 4 do
      begin
        if GetParamSocSecur('SO_PGLIBCOMBO' + IntToStr(J),'') <> '' then
        begin
          LValeur.Values.Add('LIBREPCMB' + IntToStr(J));
          LValeur.items.Add(GetParamSocSecur('SO_PGLIBCOMBO' + IntToStr(J),''));
        end
        else
        begin
          LValeur.Values.Add('LIBREPCMB' + IntToStr(J));
          LValeur.items.Add('Libellé tablette ' + IntToStr(J));
        end;
      end;

      // Zones dates libres
      for J := 1 to 4 do
      begin
        if GetParamSocSecur('SO_PGLIBDATE' + IntToStr(J),'') <> '' then
        begin
          LValeur.Values.Add('DATELIBRE' + IntToStr(J));
          LValeur.items.Add(GetParamSocSecur('SO_PGLIBDATE' + IntToStr(J),''));
        end
        else
        begin
          LValeur.Values.Add('DATELIBRE' + IntToStr(J));
          LValeur.items.Add('Libellé date libre ' + IntToStr(J));
        end;
      end;

      // Zones boites à cocher libres
      for J := 1 to 4 do
      begin
        if GetParamSocSecur('SO_PGLIBCOCHE' + IntToStr(J),'') <> '' then
        begin
          LValeur.Values.Add('BOOLLIBRE' + IntToStr(J));
          LValeur.items.Add(GetParamSocSecur('SO_PGLIBCOCHE' + IntToStr(J),''));
        end
        else
        begin
          LValeur.Values.Add('BOOLLIBRE' + IntToStr(J));
          LValeur.items.Add('Libellé boite à cocher ' + IntToStr(J));
        end;
      end;
    end;
  end;
end;

procedure TOF_PGPARAMETRES.ChargementTabZonesSalaries();
var
  I,J : integer;
  LValeur : THValComboBox;
begin
  for I := 1 to 24 do
  begin
    LValeur := THValComboBox(GetControl('ZSAL' + IntToStr(I)));
    if LValeur <> nil then
    begin

{
Coefficient (affiche libellé de la tablette)      PSA_COEFFICIENT
Qualification (affiche libellé de la tablette)    PSA_QUALIFICATION
Emploi PCS                                        PSA_CODEEMPLOI
Emploi (libellé)                                  PSA_LIBELLEEMPLOI
Condition d'emploi                                PSA_CONDEMPLOI
Date d'ancieneté                                  PSA_DATEANCIENNETE
Date d'entrée                                     PSA_DATEENTREE
Horaire                                           PSA_HORAIREMOIS
Salaire                                           PSA_SALAIRETHEO
Taux horaire                                      PSA_TAUXHORAIRE
Salaire mensuel                                   PSA_SALAIREMOIS1
Salaire annuel                                    PSA_SALAIRANN1
Organisation 1 ( libellé paramétré dans paramètre société )          PSA_TRAVAILN1    SO_PGLIBORGSTAT1
Organisation 2 idem                                                  PSA_TRAVAILN2    SO_PGLIBORGSTAT2
Organisation 3 idem                                                  PSA_TRAVAILN3    SO_PGLIBORGSTAT3
Organisation 4 idem                                                  PSA_TRAVAILN4    SO_PGLIBORGSTAT4
Indice    (affiche libellé de la tablette)         PSA_INDICE
Niveau   (affiche libellé de la tablette)          PSA_NIVEAU
Statut professionnel DADSU (libellé)               PSA_DADSPROF
Statut catégoriel DADSU      (libellé)             PSA_DADSCAT
Catégorie DUCS                  (libellé)          PSA_CATDADS
Catégorie bilan social           (libellé)         PSA_CATBILAN
Convention collective            (libellé)         PSA_CONVENTION
}

      LValeur.Values.Add('COEFFICIENT');
      LValeur.items.Add('Coefficient');
      LValeur.Values.Add('QUALIFICATION');
      LValeur.items.Add('Qualification');
      LValeur.Values.Add('CODEEMPLOI');
      LValeur.items.Add('Emploi PCS');
      LValeur.Values.Add('LIBELLEEMPLOI');
      LValeur.items.Add('Emploi');
      LValeur.Values.Add('CONDEMPLOI');
{PT10
      LValeur.items.Add('Condition d''emploi');
}
      LValeur.items.Add('Caractéristique activité');
//FIN PT10
      LValeur.Values.Add('DATEANCIENNETE');
      LValeur.items.Add('Date d''ancienneté');
      LValeur.Values.Add('DATEENTREE');
      LValeur.items.Add('Date d''entrée');
      LValeur.Values.Add('HORAIREMOIS');
      LValeur.items.Add('Horaire');
      LValeur.Values.Add('SALAIRETHEO');
      LValeur.items.Add('Salaire');
      LValeur.Values.Add('TAUXHORAIRE');
      LValeur.items.Add('Taux horaire');
      LValeur.Values.Add('SALAIREMOIS1');
      LValeur.items.Add('Salaire mensuel');
      LValeur.Values.Add('SALAIRANN1');
      LValeur.items.Add('Salaire annuel');
      LValeur.Values.Add('INDICE');
      LValeur.items.Add('Indice');
      LValeur.Values.Add('NIVEAU');
      LValeur.items.Add('Niveau');
      LValeur.Values.Add('DADSPROF');
      LValeur.items.Add('Statut professionnel DADSU');
      LValeur.Values.Add('DADSCAT');
      LValeur.items.Add('Catégorie DUCS');
      LValeur.Values.Add('CATBILAN');
      LValeur.items.Add('Catégorie bilan social');
      LValeur.Values.Add('CONVENTION');
      LValeur.items.Add('Convention collective');

      // Zones organisation
      for J := 1 to 4 do
      begin
        if GetParamSocSecur('SO_PGLIBORGSTAT' + IntToStr(J),'') <> '' then
        begin
          LValeur.Values.Add('TRAVAILN' + IntToStr(J));
          LValeur.items.Add(GetParamSocSecur('SO_PGLIBORGSTAT' + IntToStr(J),''));
        end
        else
        begin
          LValeur.Values.Add('TRAVAILN' + IntToStr(J));
          LValeur.items.Add('Libellé organisation ' + IntToStr(J));
        end;
      end;
    end;
  end;
end;

Procedure TOF_PGPARAMETRES.CheckValuesBUL(Sender: TObject);
var
  LValeur , ComboType, ComboValeur, ComboValeurB, ComboTest :THValComboBox;
  Valeur :THEdit;
  intJ :Integer;
  Name, TypeSender, stI, strJ, strDataType :String;
  procedure VerificationZoneResultat(ComboTest:THValComboBox;stI:String;strJ:String);
  begin
    if (ComboTest.Value = 'ETQ') or (ComboTest.Value = 'ELD') or (ComboTest.Value = 'RES') or (ComboTest.Value = 'SAL') or (ComboTest.Value = 'ZLS') then //PT2  //PT4
      LastErrorMsg := TraduireMemoire('La Zone ') + stI
                    + TraduireMemoire(' de type résultat'#10#13'doit utiliser des zones de type autre que Etiquette, Elément dynamique libre, Donnée fiche salarié, Zones libres et Résultat.');
    if (ComboTest.Value = '') then
      LastErrorMsg := TraduireMemoire('La Zone ') + stI
                    + TraduireMemoire(' de type Résultat'#10#13'doit utiliser une zone avec un type renseigné.');
    if (stI = strJ) then
      LastErrorMsg := TraduireMemoire('La Zone ') + stI
                    + TraduireMemoire(' de type Résultat'#10#13'ne peut s''appeler elle même.');
  end;
begin
  LastErrorMsg := '';
  if Sender is TControl then
  begin
    Name := (Sender as TControl).Name;
    //On récupère l'indice du control pour tester la bonne ligne
    if Pos('CBVALZONE',Name) > 0 then
    begin
      TypeSender := 'CBVALZONE';
      stI := RightStr(Name,Length(Name)-9);
      if Pos('B',stI) > 0 then
      begin
        TypeSender := 'CBVALZONEB';
        stI := LeftStr(stI, Length(stI)-1);
      end;
    end else if Pos('VALZONE',Name) > 0 then
    begin
      TypeSender := 'VALZONE';
      stI := RightStr(Name,Length(Name)-7);
    end;
  end else exit;

  ComboType := THValComboBox(GetControl('TYPZONE' + stI));
  if (ComboType <> nil) and (ComboType.Value <> '') then
  begin
    ComboValeur    := THValComboBox(GetControl('CBVALZONE' + stI));
    if TypeSender = 'CBVALZONE' then
    begin
      // Controles pour les zones type Résultat
      if (ComboType.Value = 'RES') and (ComboValeur.Value <> '') then
      begin
        // Controle zone à gauche de l'opérateur
        strJ := ComboValeur.Value; intJ := StrToInt(strJ); strJ := Inttostr(intJ); // Suppression de l'éventuel zéro de gauche
        ComboTest := THValComboBox(GetControl('TYPZONE' + strJ));
        VerificationZoneResultat(ComboTest,stI,strJ);
      end;
      (GetControl('PAGES') as TPageControl).ActivePageIndex := 5 + (StrToInt(stI) div 13);
      if LastErrorMsg <> '' then (GetControl('CBVALZONE' + stI) as THValComboBox).SetFocus;
    end else if TypeSender = 'CBVALZONEB' then
    begin
      // Controles pour les zones type Résultat
      if (ComboType.Value = 'RES') and (ComboValeur.Value <> '') then
      begin
        ComboValeurB   := THValComboBox(GetControl('CBVALZONE' + stI + 'B'));
        // Controle zone à droite de l'opérateur
        strJ := ComboValeurB.Value; intJ := StrToInt(strJ); strJ := Inttostr(intJ); // Suppression de l'éventuel zéro de gauche
        ComboTest := THValComboBox(GetControl('TYPZONE' + strJ));
        VerificationZoneResultat(ComboTest,stI,strJ);
      end;
      (GetControl('PAGES') as TPageControl).ActivePageIndex := 5 + (StrToInt(stI) div 13);
      if LastErrorMsg <> '' then (GetControl('CBVALZONE' + stI+ 'B') as THValComboBox).SetFocus;
    end else if TypeSender = 'VALZONE' then
    begin
      Valeur         := THEdit(GetControl('VALZONE' + stI));
      // Controles pour les zones autre que type Résultat
      if (ComboType.Value <> 'RES') and (ComboType.Value <> 'ZLS') and (ComboType.Value <> 'SAL') and (Valeur.Text = '') then
        LastErrorMsg := TraduireMemoire('La Zone ')+stI+ TraduireMemoire(' de type')+' '
                      + ComboType.Text + TraduireMemoire(' n''a pas de valeur renseignée.');
      if ((ComboType.Value <> 'RES') and (Valeur.Text <> '')) or (ComboType.value = 'ZLS') or (ComboType.value = 'SAL') then
      begin
        if (ComboType.value = 'CMP') or (ComboType.value = 'CMM') then
          strDataType := 'PGCUMULPAIE';
        if (ComboType.value = 'BAS') or (ComboType.value = 'TAU') or (ComboType.value = 'COE') or (ComboType.value = 'MNT') then
          strDataType := 'PGREMUNERATION';
        if (ComboType.value = 'ELD') then
          strDataType := 'PGZONEELTDYN';
        if (ComboType.value = 'ZLS') then
        begin
          LValeur := THValComboBox(GetControl('ZLS' + stI));
          if LValeur <> nil then
          begin
            if (LValeur.Items.ValueFromIndex[LValeur.ItemIndex] = '') then
              LastErrorMsg := TraduireMemoire('La Zone ')+stI+ TraduireMemoire(' de type')+' '
                      + ComboType.Text + TraduireMemoire(' n''a pas de valeur renseignée.');
          end;
        end;
        if (ComboType.value = 'ELM') then //PT2
          strDataType := 'PGELEMENTNAT';
        if (ComboType.value = 'SAL') then
        begin
          LValeur := THValComboBox(GetControl('ZSAL' + stI));
          if LValeur <> nil then
          begin
            if (LValeur.Items.ValueFromIndex[LValeur.ItemIndex] = '') then
              LastErrorMsg := TraduireMemoire('La Zone ')+stI+ TraduireMemoire(' de type')+' '
                      + ComboType.Text + TraduireMemoire(' n''a pas de valeur renseignée.');
          end;
        end;
        if strDataType <> '' then
          if Rechdom(strDataType, Valeur.Text, False) = '' then
            LastErrorMsg := TraduireMemoire('Zone ')+ stI +#10#13+TraduireMemoire('La valeur saisie n''existe pas.');
      end;
      (GetControl('PAGES') as TPageControl).ActivePageIndex := 5 + (StrToInt(stI) div 13);
      if LastErrorMsg <> '' then (GetControl('VALZONE' + stI) as THEdit).SetFocus;
    end;
  end;
  if LastErrorMsg <> '' then LastError := 1;
  CheckError;
end;

procedure TOF_PGPARAMETRES.BtnAddClick(Sender: TObject);
begin
  GD_GRILLEFRAC.RowCount := GD_GRILLEFRAC.RowCount + 5;
  GD_GRILLEFRAC.Col := 0;
end;

procedure TOF_PGPARAMETRES.BtnDelClick(Sender: TObject);
var
  indexCol : Integer;

  Procedure  DeleteRowGridColonne(ColName : String; Row : Integer);
  var
    TempTob : TOB;
    indexRow : Integer;
  begin
    indexRow := Row;
    { Suppression de la ligne sélectionnée }
    TempTob := TobParametres.FindFirst( ['PGP_PGPARAMETRE','PGP_TYPEPARAMETRE','PGP_PREDEFINI','PGP_NODOSSIER','PGP_MODULECHAMP','PGP_PGNOMCHAMP']
                                      ,[init_PGPARAMETRE,init_TYPEPARAMETRE,init_PREDEFINI,GetNoDossier(init_PREDEFINI),init_MODULECHAMP,ColName+';'+intToStr(indexRow)]
                                      ,True);
    if Assigned(TempTob) then
    begin
      TempTob.ChangeParent(nil,-1);
      FreeAndNil(TempTob);
    end;
    { Décalage des lignes suivantes }
    Inc(indexRow);
    TempTob := TobParametres.FindFirst(['PGP_PGPARAMETRE','PGP_TYPEPARAMETRE','PGP_PREDEFINI','PGP_NODOSSIER','PGP_MODULECHAMP','PGP_PGNOMCHAMP']
                                     ,[init_PGPARAMETRE,init_TYPEPARAMETRE,init_PREDEFINI,GetNoDossier(init_PREDEFINI),init_MODULECHAMP,ColName+';'+intToStr(indexRow)]
                                     ,True);
    while Assigned(TempTob) do
    begin
      TempTob.PutValue('PGP_PGNOMCHAMP', ColName+';'+intToStr(indexRow-1));
      Inc(indexRow);
      TempTob := TobParametres.FindFirst(['PGP_PGPARAMETRE','PGP_TYPEPARAMETRE','PGP_PREDEFINI','PGP_NODOSSIER','PGP_MODULECHAMP','PGP_PGNOMCHAMP']
                                       ,[init_PGPARAMETRE,init_TYPEPARAMETRE,init_PREDEFINI,GetNoDossier(init_PREDEFINI),init_MODULECHAMP,ColName+';'+intToStr(indexRow)]
                                       ,True);
    end;
  end;

begin
  if GD_GRILLEFRAC.Row <> 0 then
  begin
    GD_GRILLEFRAC.DeleteRow(GD_GRILLEFRAC.Row);
    for indexCol := 0 to GD_GRILLEFRAC.ColCount -1 do
    begin
      DeleteRowGridColonne(GD_GRILLEFRAC.ColNames[indexCol],GD_GRILLEFRAC.Row);
    end;
  end;
end;

procedure TOF_PGPARAMETRES.BtnInsClick(Sender: TObject);
begin
  GD_GRILLEFRAC.InsertRow(GD_GRILLEFRAC.Row);
end;

procedure TOF_PGPARAMETRES.ShowLookupPROFILPRES(Sender: TObject); // PT6
begin
    LookupList(GetControl('PROFILPRES')
              ,'Liste des paramètres'
              ,'PROFILPRESENCE a'
              ,'PPQ_PROFILPRES'
              ,'PPQ_LIBELLE'
              ,'a.PPQ_DATEVALIDITE = (Select MAX(b.PPQ_DATEVALIDITE) from profilpresence b where a.PPQ_PROFILPRES = b.PPQ_PROFILPRES) '
              ,''
              ,True
              ,-1
            //  ,'select PPQ_PROFILPRES, PPQ_LIBELLE from PROFILPRESENCE a where a.PPQ_DATEVALIDITE = (Select ##TOP 1## b.PPQ_DATEVALIDITE from profilpresence b where a.PPQ_PROFILPRES = b.PPQ_PROFILPRES  order by b.PPQ_DATEVALIDITE desc) '
            );
end;

Initialization
  registerclasses ( [ TOF_PGPARAMETRES ] ) ;
end.
