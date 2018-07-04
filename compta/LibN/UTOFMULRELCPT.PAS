{***********UNITE*************************************************
Auteur  ...... : BOUSSERT Stéphane
Créé le ...... : 25/04/2003
Modifié le ... : 25/04/2003
Description .. : Source TOF de la FICHE : CPMULRELCPT ()
Mots clefs ... : TOF;CPMULRELCPT
*****************************************************************}
unit UTOFMULRELCPT;

interface

uses StdCtrls,
  Controls,
  Classes,
  Menus, // TPopupMenu
  CPTIERS_TOM, // pour Zoom sur fiche tiers
  {$IFDEF EAGLCLIENT}
  UTob,
  maineAGL,
  UtileAGL,
  eMul,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  fe_main,
  EdtRDoc,
  EdtDoc,
  Mul,
  hdb,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HTB97,
  HEnt1,
  Ent1,
  Ed_Tools,
  HMsgBox,
  HQry,
  HStatus, // pour Initmove
  LettUtil, // pour InitTablesLibresTiers
  ZECRIMVT_TOF, // pour le zoom sur les pièces
  uRecupSQLModele, // pout RecupSQLModele
  filtre,
  ParamSoc,		// GetParamSocSecur YMO
  UTOF;




procedure ReleveCompte;
procedure CPLanceFiche_MulRelCpt(vStRange, vStLequel, vStArgs: string);


type
  TOF_CPMULRELCPT = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure OnAfterFormShow          ; // FQ18694
  private
    gLoad : boolean ; // FQ 21838
    Q: THQuery;
    {$IFDEF EAGLCLIENT}
    FListe: THGrid;
    {$ELSE}
    FListe: THDBGrid;
    {$ENDIF}
    EXERCICE: THValComboBox;
    ORDREEDIT: THValComboBox;
    MODELE: THValComboBox;
    DEVISE: THValComboBox;
    AUXILIAIRE: THEdit;
    AUXILIAIRE_: THEdit;
    COLLECTIF: THEdit;
    COLLECTIF_: THEdit;
    DATECOMPTABLE: THEdit;
    DATECOMPTABLE_: THEdit;
    CHOIXNAT: THRadioGroup;
    AVECLETTRAGE: TCheckBox;
    AVECANOUV: TCheckBox;
    APP: TCheckBox;
    ETABLISSEMENT: THMultiValComboBox;
    HM: THMsgBox;
    // popup Menu
    POPZ: TPopUpMenu;
    bSelectAllClickParent     : TNotifyEvent ;

    // Initialisations
    procedure InitComposants;
    procedure InitEvenements;
    procedure InitMsgBox;
    // Boutons
    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BParamModeleClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BMenuZoomClick(Sender: TObject);
    // Filtres
    procedure ChoixNatExit(Sender: TObject);
    procedure EXERCICEChange(Sender: TObject);
    // Utils
    function JeValide: boolean;
    function FabricSQLEcr: string;
    function OkMvt(SQL: string): boolean;
    procedure EditeRelevesCompte;
    function  OrdreTri : String ;  // FQ 19618
    procedure bSelectAllClick(Sender: TObject);

    // Gestion des filtres
    procedure NouvRechClick      (Sender: TObject) ; // FQ18694
    procedure AuxiElipsisClick(Sender : TObject);

  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}

  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }


// =============================================================================

procedure ReleveCompte;
begin
  CPLanceFiche_MulRelCpt('', '', '');
end;

// =============================================================================

procedure CPLanceFiche_MulRelCpt(vStRange, vStLequel, vStArgs: string);
begin
  AGLLanceFiche('CP', 'CPMULRELCPT', vStRange, vStLequel, vStArgs);
end;

// =============================================================================

procedure TOF_CPMULRELCPT.OnNew;
begin
  inherited;
end;

procedure TOF_CPMULRELCPT.OnDelete;
begin
  inherited;
end;

procedure TOF_CPMULRELCPT.OnUpdate;
begin
  inherited;
  { FQ 21838 BVE 22.11.2007 }
  if not gLoad then OnLoad;
  gLoad := false;
  { END FQ 21838 }
end;

procedure TOF_CPMULRELCPT.OnLoad;
var lStWhere : string ;
OkTri,OkNumCol : boolean;
NewValue,Table,Crit,Tri,Champ,Larg,Align,Params,Perso : String;
Titre,LeTitre,NomCol : WideString;
begin
  inherited;
  EXERCICE.Vide := true;   
  { FQ 21838 BVE 22.11.2007 }
  gLoad := true;   
  { END FQ 21838 }

  // FQ18693
  lStWhere := GetControlText('XX_WHERE') ;
  { FQ 19618 BVE 24.04.07 }
  if not (GetCheckBoxState('AVECLETTRAGE') = cbChecked) then
    lStWhere := '(' + lStWhere + ') AND E_ETATLETTRAGE<>"TL" ';
  if Pos('E_MODEPAIE <> ""',lStWhere) = 0 then             
    lStWhere := '(' + lStWhere + ') AND E_MODEPAIE <> "" ';
  SetControlText('XX_WHERE', '(' + lStWhere + ') ' ) ;
  NewValue := 'RELEVECOMPTE';
  ChargeHListe(NewValue,Table,Crit,Tri,Champ,Titre,Larg,Align,Params,LeTitre,NomCol,Perso,OkTri,OkNumCol);
  Champ := StringReplace(Champ,';',',',[rfReplaceAll]);
  System.Delete(Champ,Length(Champ),1);
  SetControlText('XX_GROUPBY',Champ);
  SetControlText('XX_ORDERBY',OrdreTri);
  { END FQ 19618 }
end;

procedure TOF_CPMULRELCPT.OnArgument(S: string);
begin
  inherited;
  InitComposants;
  InitEvenements;
  InitMsgBox;


  {$IFDEF CCMP}
  if (VH^.CCMP.LotCli) then
  begin
    CHOIXNAT.ItemIndex := 0;
    CHOIXNAT.Visible := False;
    SetControlVisible('TChoixNat', False);
    // BPY le 09/09/2004 : Fiche n° 14522 : set au bon type d'auxiliaire !
    SetControlText('XX_WHERE', 'T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD"');
    // Fin BPY
  end
  else
  begin
    CHOIXNAT.ItemIndex := 1;
    CHOIXNAT.Visible := False;
    SetControlVisible('TChoixNat', False);
    // BPY le 09/09/2004 : Fiche n° 14522 : set au bon type d'auxiliaire !
    SetControlText('XX_WHERE', 'T_NATUREAUXI="FOU" or T_NATUREAUXI="AUC"');
    // Fin BPY
  end;
  {$ENDIF}

  //Modification proprietés combo exercice


  if VH^.CPExoRef.Code <> '' then
  begin
    EXERCICE.Value := VH^.CPExoRef.Code;
    DATECOMPTABLE.Text := DateToStr(VH^.CPExoRef.Deb);
    DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin);
  end
  else
  begin
    EXERCICE.Value := VH^.Entree.Code;
    DATECOMPTABLE.Text := StDate1900;
    DATECOMPTABLE_.Text := StDate2099;
  end;

  InitTablesLibresTiers(TTabSheet(GetControl('PLibres', True)));
  DEVISE.Value := V_PGI.DevisePivot;
  ORDREEDIT.ItemIndex := 0;
  if MODELE.Values.Count > 0
    then MODELE.Value := MODELE.Values[0];

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    AUXILIAIRE.OnElipsisClick:=AuxiElipsisClick;
    AUXILIAIRE_.OnElipsisClick:=AuxiElipsisClick;
  end;

  { FQ 19618 BVE 17.09.07 }
  SetControlVisible('bSelectAll',true);
  { END FQ 19618 }

  {JP 30/10/07 : FQ 21432 : Glyph du bouton valider "VVVert"}
  TFMul(Ecran).BOuvrir.GlobalIndexImage := 'Z0184_S16G1';
  { FQ 21838 BVE 22.11.2007 }
  gLoad := false;   
  { END FQ 21838 }
end;

procedure TOF_CPMULRELCPT.OnClose;
begin
  inherited;
end;

procedure TOF_CPMULRELCPT.OnDisplay();
begin
  inherited;
end;

procedure TOF_CPMULRELCPT.OnCancel();
begin
  inherited;
end;

procedure TOF_CPMULRELCPT.BOuvrirClick(Sender: TObject);
begin
  if not JeValide then Exit;
  Application.ProcessMessages;
  EditeRelevesCompte;
  FListe.ClearSelected;
  {YMO 14/06/2006 FQ18104 Rafraîchissement de la liste pour le bselectallclick}
  TFMul(Ecran).BCherche.Click;
end;

procedure TOF_CPMULRELCPT.BParamModeleClick(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
  // Non compatible eAGL
  {$ELSE}
  EditDocument('L', 'RLC', MODELE.Value, True);
  {$ENDIF}
end;

procedure TOF_CPMULRELCPT.BZoomClick(Sender: TObject);
var
  Cpte: string;
begin
  inherited;
  {$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then Exit;
  Q.TQ.Seek(FListe.row - 1);
  Cpte := Q.FindField('T_AUXILIAIRE').AsString;
  FicheTiers(nil, '', Cpte, taConsult, 1);
  {$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit;
  Cpte := Q.FindField('T_AUXILIAIRE').AsString;
  FicheTiers(Q, '', Cpte, taConsult, 1);
  {$ENDIF}

end;

procedure TOF_CPMULRELCPT.BZoomPieceClick(Sender: TObject);
var
  Auxi: string;
begin
  inherited;
  {$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then Exit;
  Q.TQ.Seek(FListe.row - 1);
  {$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit;
  {$ENDIF}
  Auxi := Q.FindField('T_AUXILIAIRE').AsString;
  ZoomEcritureMvt(Auxi, fbAux, 'MULMMVTS');
end;

procedure TOF_CPMULRELCPT.ChoixNatExit(Sender: TObject);
begin
  if ChoixNat.ItemIndex = 0
    then SetControlText('XX_WHERE', 'T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD"')
  else SetControlText('XX_WHERE', 'T_NATUREAUXI="FOU" or T_NATUREAUXI="AUC"');
end;

//YMO 24/04/2006 : EXERCICEExit est systématiquement déclenché lors de la recherche
procedure TOF_CPMULRELCPT.EXERCICEChange(Sender: TObject);
begin
  ExoToDates(EXERCICE.Value, DATECOMPTABLE, DATECOMPTABLE_);
  if EXERCICE.Value = '' then
  begin
    DATECOMPTABLE.Text := StDate1900;
    DATECOMPTABLE_.Text := StDate2099;
  end;
end;

procedure TOF_CPMULRELCPT.FListeDblClick(Sender: TObject);
begin
  BZoomPieceClick(nil);
end;

procedure TOF_CPMULRELCPT.InitComposants;
begin
  Q := THQuery(GetControl('Q', True));
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe', True));
  {$ELSE}
  FListe := THDBGrid(GetControl('FListe', True));
  {$ENDIF}
  // Onglet "Standards"
  ETABLISSEMENT := THMultiValComboBox(GetControl('E_ETABLISSEMENT', True)); // FQ 19618
  DEVISE := THValComboBox(GetControl('E_DEVISE', True));// FQ 19618
  AUXILIAIRE := THEdit(GetControl('T_AUXILIAIRE', True));
  AUXILIAIRE_ := THEdit(GetControl('T_AUXILIAIRE_', True));
  // YMO 06/06/2006 Correction plantage au lancement
  COLLECTIF := THEdit(GetControl('T_COLLECTIF', True));
  COLLECTIF_ := THEdit(GetControl('T_COLLECTIF_', True));
  CHOIXNAT := THRadioGroup(GetControl('CHOIXNAT', True));
  // Onglet "Ecritures et edition"
  EXERCICE := THValComboBox(GetControl('E_EXERCICE', True)); // FQ 19618
  DATECOMPTABLE := THEdit(GetControl('DATECOMPTABLE', True));
  DATECOMPTABLE_ := THEdit(GetControl('DATECOMPTABLE_', True));
  AVECLETTRAGE := TCheckBox(GetControl('AVECLETTRAGE', True));
  AVECANOUV := TCheckBox(GetControl('AVECANOUV', True));
  MODELE := THValComboBox(GetControl('MODELE', True));
  ORDREEDIT := THValComboBox(GetControl('ORDREEDIT', True));
  APP := TCheckBox(GetControl('APP', True));
  // Boutons
  {$IFDEF EAGLCLIENT}
  TButton(GetControl('BParamModele', True)).Visible := False;
  {$ENDIF}

  {b FP 07/11/2005 FQ16883}
  // PopupMenu
  //POPZ := TPopUpMenu.Create(Ecran);
  POPZ := TPopupMenu(getcontrol('POPZ',true));
  {e FP 07/11/2005 FQ16883}

  //Modification propriétés combo exercice
  EXERCICE.Vide := true;
  EXERCICE.ReLoad;
  PositionneEtabUser(ETABLISSEMENT, False); // 15088

end;

procedure TOF_CPMULRELCPT.InitEvenements;
Var
  lBouton   : TButton ;
begin
  // Grille
  {$IFDEF EAGLCLIENT}
  FListe.OnDblClick := FListeDblClick;
  {$ELSE}
  FListe.OnDblClick := FListeDblClick;
  {$ENDIF}
  // Boutons
  TButton(GetControl('BOuvrir', True)).OnClick := BOuvrirClick;
  TButton(GetControl('BParamModele', True)).OnClick := BParamModeleClick;
  // menuZoom
  TToolBarButton97(GetControl('BMenuZoom', true)).OnClick := BMenuZoomClick;
  TButton(GetControl('BZOOM', True)).OnClick := BZoomClick;
  TButton(GetControl('BZoomPiece', True)).OnClick := BZoomPieceClick;
  // Filtres
  CHOIXNAT.OnExit := ChoixNatExit;
  EXERCICE.OnChange := EXERCICEChange;

 // -> tout sélectionner
  lBouton := TButton(GetControl('bSelectAll',True)) ;
  if Assigned(lBouton.OnClick)
    then bSelectAllClickParent    := lBouton.OnClick
    else bSelectAllClickParent    := nil ;
  lBouton.OnClick := bSelectAllClick ;

  TFMul(Ecran).OnAfterFormShow := OnAfterFormShow; // FQ18694

end;

procedure TOF_CPMULRELCPT.EditeRelevesCompte;
var
  i, j, posA: integer;
  WhereEcr, Sql, StOrder, SQLWhere : string;
  TT: TList;
  LAuxi, LModeP : TStringList;
  LL: TStrings;

begin
  WhereEcr := FabricSQLEcr;

  InitMove(FListe.NbSelected, '');
  LAuxi:=TStringList.Create;
  LModeP:=TStringList.Create;

  // YMO 14/06/2006 FQ18104 Rajout bouton 'tout sélectionner'
  if TFmul(Ecran).FListe.AllSelected then
  begin
      TFMul(Ecran).Q.First;
      While not TFMul(Ecran).Q.EOF do
      begin

        posA := LAuxi.IndexOf(Q.FindField('T_AUXILIAIRE').AsString);

        If posA<0 then
        begin
          LAuxi.Add(Q.FindField('T_AUXILIAIRE').AsString);
          LModeP.Add('"'+Q.FindField('E_MODEPAIE').AsString+'"');
        end
        else
        LModeP.Strings[posA]:= LModeP.Strings[posA]+',"'+Q.FindField('E_MODEPAIE').AsString+'"';

        TFMul(Ecran).Q.Next;
      end;
  end
  else
  for i := 0 to FListe.NbSelected - 1 do
  begin
    MoveCur(FALSE);
    FListe.GotoLeBookMark(i);
    {$IFDEF EAGLCLIENT}
    Q.TQ.Seek(FListe.Row - 1);
    {$ENDIF}

    posA := LAuxi.IndexOf(Q.FindField('T_AUXILIAIRE').AsString);

    If posA<0 then
    begin
      LAuxi.Add(Q.FindField('T_AUXILIAIRE').AsString);
      LModeP.Add('"'+Q.FindField('E_MODEPAIE').AsString+'"');
    end
    else
      LModeP.Strings[posA]:= LModeP.Strings[posA]+',"'+Q.FindField('E_MODEPAIE').AsString+'"';

  end;

  TT := Tlist.Create;

  for i := 0 to LAuxi.Count-1 do
  begin

    // YMO 25/04/2006 On ne prend que le mode de paiement choisi
    LL := TStringList.Create;

    SQLWhere:=' WHERE T_AUXILIAIRE="'+LAuxi.Strings[i]+'" AND ((E_MODEPAIE IN ('+LModeP.Strings[i]+'))';

     // YMO 14/06/2006 FQ18104 Choix de prendre les a-nouveaux ou pas
     // CA 25/07/2006 - Ne pas prendre les à nouveaux à ce niveau
(*
    if (GetCheckBoxState('AvecANouv') = cbChecked) then
    begin
      if GetControltext('EXERCICE') <> '' then
         SQLWhere := SQLWhere +' OR E_ECRANOUVEAU="OAN")' //YMO 24/05/2006 Rajout des A-nouveaux
      else
         SQLWhere := SQLWhere +' OR E_ECRANOUVEAU="H")' //Si on prend tous les exos, Uniqut RAN de type 'H'
    end
    else
*)
    SQLWhere := SQLWhere+')';

    Sql := RecupSQLModele('L', 'RLC', Modele.Value, '', '', '', SQLWhere);


    StOrder := '';
    j := Pos('ORDER BY', Sql);

    if j > 0 then
    begin
      StOrder := Copy(Sql, j, Length(Sql) - j + 1);
      System.Delete(Sql, j, Length(Sql) - j + 1);
{ FQ 19618 BVE 24.04.07 }
      {JP 23/10/07 : Plantage si Order By est renseigné}
      if (Trim(StOrder) <> '') and (Trim(StOrder)[Length(Trim(StOrder))] <> ',') then
        StOrder := StOrder + ',';

      StOrder := StOrder + ' ' + OrdreTri;
    end
    else if OrdreTri <> '' then {JP 23/10/07 : Sur Ctrl + Z, cela peut renvoyer vide}
      StOrder := 'ORDER BY ' + OrdreTri ;
{ END FQ 19618 }

    if Trim(WhereEcr) <> '' then Sql := Sql + ' AND ' + WhereEcr;
    if StOrder <> '' then Sql := Sql + ' ' + StOrder;

    if OkMvt(SQL) then
    begin
      LL.Add(Sql);
      TT.Add(LL);
    end
    else LL.Free;

  end;

  FiniMove;

  {JP 23/10/07 : FQ 21715 : pour éviter que l'état mouline à vide en eAGL}
  if TT.Count > 0 then
    LanceDocument('L', 'RLC', MODELE.Value, TT, nil, APP.Checked, True)
  else
    PGIInfo(TraduireMemoire('Le document ne renvoie pas de données'), Ecran.Caption);

  VideListe(TT);
  TT.Free;
end;

function TOF_CPMULRELCPT.FabricSQLEcr: string;
var
  St, St2, StOr, St3, StV8 : string;
begin
  St := 'E_DATECOMPTABLE>="' + UsDateTime(StrToDate(GetControlText('DATECOMPTABLE'))) + '" '
    + 'AND E_DATECOMPTABLE<="' + UsDateTime(StrToDate(GetControlText('DATECOMPTABLE_'))) + '" '
    + 'AND E_DEVISE="' + GetControlText('E_DEVISE') + '" ' // FQ 19618
    + 'AND E_QUALIFPIECE="N" AND E_TRESOLETTRE<>"X" ' ;
  // YMO 24/05/2006 Rajout selection sur cptes collectifs
  // YMO 30/05/2006 Mise en place avec SOCREF 749
  if COLLECTIF.text <>'' then St := St + 'AND E_GENERAL>="' + COLLECTIF.text + '" ';
  if COLLECTIF_.text<>'' then St := St + 'AND E_GENERAL<="' + COLLECTIF_.text + '" ';

  if GetControltext('E_EXERCICE') <> '' // FQ 19618
    then St := St + ' AND E_EXERCICE="' + GetControltext('E_EXERCICE') + '"'; // FQ 19618
  St := St + ' AND ((E_FLAGECR<>"ROR" AND E_FLAGECR<>"RDE") OR (E_FLAGECR="")) ';
  if (GetControltext('E_ETABLISSEMENT') <> '') and (GetControltext('E_ETABLISSEMENT') <> Traduirememoire('<<Tous>>')) then // FQ 19618
  begin
    St2 := GetControltext('E_ETABLISSEMENT'); // FQ 19618
    StOr := '';
    while St2 <> '' do
      StOr := StOr + ' OR E_ETABLISSEMENT="' + ReadTokenSt(St2) + '"';
    System.Delete(StOr, 1, 4);
    if StOR <> '' then
      St := St + ' AND (' + StOr + ')';
  end;
  //  if GetControltext('SAISIEEURO')<>'' then
    //  St := St + ' AND E_SAISIEEURO="' + GetControltext('SAISIEEURO') + '"' ;
  if not (GetCheckBoxState('AvecLettrage') = cbChecked) then
    St := St + ' AND E_ETATLETTRAGE<>"TL" ';
  St3 := ' (E_ECRANOUVEAU="N"';

  // YMO 14/06/2006 FQ18104 Choix de prendre les a-nouveaux ou pas
  if (GetCheckBoxState('AvecANouv')=cbChecked) and  (GetControltext('E_EXERCICE')<>'') then
    St3 := St3 + ' OR E_ECRANOUVEAU="OAN" '
  else St3 := St3+'';

  StV8 := LWhereV8;
  StV8:=LWhereV8 ;
  if StV8<>'' then
  BEGIN
    St3:='('+St3+' OR E_ECRANOUVEAU="H"))  ' ;
    St3:=St3+' AND ('+StV8+') ' ;
  end else St3 := St3+')';


  St := St + ' AND ' + St3;
  Result := St;
end;

function TOF_CPMULRELCPT.JeValide: boolean;
var
  ii: integer;
begin
  Result := False;
  {$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then
    {$ELSE}
  if ((Q.EOF) and (Q.BOF)) then
    {$ENDIF}
  begin
  //YMO 14/06/2006 Message plus approprié
    HM.Execute(5, '', '');
    Exit;
  end;
  if (FListe.NbSelected <= 0) and (not FListe.allSelected) then
  begin
    HM.Execute(0, '', '');
    Exit;
  end;
  if MODELE.Value = '' then
  begin
    HM.Execute(1, '', '');
    Exit;
  end;
  ii := HM.Execute(2, '', '');
  case ii of
    mrNo:
      begin
        FListe.ClearSelected;
        Exit;
      end;
    mrCancel: Exit;
  end;
  Result := True;
end;

function TOF_CPMULRELCPT.OkMvt(SQL: string): boolean;
begin
  Result := ExisteSQL(SQL);
end;

{ FQ 19618 BVE 24.04.07 }
function TOF_CPMULRELCPT.OrdreTri: String;
begin
  Result:='' ;
  if ORDREEDIT.Value='CPT'
    then Result:=' T_AUXILIAIRE,T_LIBELLE'
    else if ORDREEDIT.Value='INT'
      then Result:=' T_LIBELLE,T_AUXILIAIRE'
      else if ORDREEDIT.Value='PCP' then
        Result:=' T_PAYS, T_CODEPOSTAL, T_AUXILIAIRE' ;
end;
{ END FQ 19618 }

procedure TOF_CPMULRELCPT.InitMsgBox;
begin
  HM := THMsgBox.create(FMenuG);
  HM.Mess.Clear;
  HM.Mess.Add('0;Relevés de compte;Vous n''avez sélectionné aucun tiers.;E;O;O;O;');
  HM.Mess.Add('1;Relevés de compte;Vous devez choisir un modèle d''édition des relevés;W;O;O;O;');
  HM.Mess.Add('2;Relevés de compte;Confirmez-vous l''édition des relevés pour les tiers sélectionnés ?;Q;YNC;Y;Y;');
  HM.Mess.Add('EURO');
  HM.Mess.Add('4');
  HM.Mess.Add('5;Le résultat de la recherche est vide. Veuillez relancer une recherche.;E;O;O;O;');
end;

procedure TOF_CPMULRELCPT.BMenuZoomClick(Sender: TObject);
begin
  PopZoom97(TToolBarButton97(GetControl('BMenuZoom', true)), POPZ);
end;

procedure TOF_CPMULRELCPT.bSelectAllClick(Sender: TObject);
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULRELCPT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPMULRELCPT.NouvRechClick(Sender: TObject);
begin

  videFiltre( THValComboBox(GetControl('FFILTRES')), TPageControl(GetControl('PAGES')) ) ;
  TFMUL(Ecran).ListeFiltre.new ;

  // === INIT ZONES
  if VH^.CPExoRef.Code <> '' then
    begin
    EXERCICE.Value      := VH^.CPExoRef.Code;
    DATECOMPTABLE.Text  := DateToStr(VH^.CPExoRef.Deb);
    DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin);
    end
  else
    begin
    EXERCICE.Value       := VH^.Entree.Code;
    DATECOMPTABLE.Text   := StDate1900;
    DATECOMPTABLE_.Text  := StDate2099;
    end;

  DEVISE.Value := V_PGI.DevisePivot;

  ORDREEDIT.ItemIndex := 0;

  if MODELE.Values.Count > 0
    then MODELE.Value := MODELE.Values[0];



end;

procedure TOF_CPMULRELCPT.OnAfterFormShow;
begin
  TFMul(Ecran).ListeFiltre.OnItemNouveau   := NouvRechClick ;
end;

initialization
  registerclasses([TOF_CPMULRELCPT]);
end.


