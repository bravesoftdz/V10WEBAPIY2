{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTJOUCON ()
Mots clefs ... : TOF;BTJOUCON
*****************************************************************}
unit BTJOUCON_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  BTPUtil,
  forms,
  sysutils,
  ComCtrls,
  Menus,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  Stat,
  HQry,
  UTOB,
  utobview,
  AGLInit,
  facture,
  uTofAfBaseCodeAffaire,
  HTB97,
  CalcOLEGenericBTP;

type
  TOF_BTJOUCON = class(TOF_AFBASECODEAFFAIRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;

  private
    TV1         : TTobViewer;
    TOBBJC      : TOB;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    CodeAffaire : String;
    CodeAffaire1: String;
    CodeAffaire2: String;
    CodeAffaire3: String;
    Ressource   : String;
    Nature      : string;
    LibAffaire  : THLabel;
    TypeAffaire : THMultiValComboBox;
    STypeaff    : String;
    BSelectAff  : TToolBarButton97;
    //
    Aff1,Aff2,Aff3,Aff4 : string;
    //
    Ok_FullScreen : Boolean;
    //FV1 : 18/02/2014 - FS#892 - BAGE : Appel journal des consommations depuis la saisie des restes à dépenser
    LanceAuto   : Boolean;
    //
    procedure OnDblClickTV(Sender: TObject);
    procedure ChargeLesConsos;
    procedure CompletelesInfos;
    procedure AddChampsSupLoc(TOBB: TOB);
    Procedure AddChampsSupLibTableLibre(TOBB: TOB);
    procedure AlimenteInfoProv(TOBL, TOBF: TOB);
    //FV1 - 29/03/2018 : FS#3025 - ECHAFAUDAGE SERVICE : Permettre d'afficher le libellé des tables libres dans le journal des conso    
    Procedure AlimenteLibelleTableLibreRessource(TOBL : TOB);
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
    procedure OnChangeTypeAffaire(Sender: TObject);
    Procedure OnClickSelectAff(Sender : TObject);
  end;

implementation
uses Windows,vierge, affaireutil;

procedure TOF_BTJOUCON.OnNew;
begin
  inherited;
end;

procedure TOF_BTJOUCON.OnDelete;
begin
  inherited;
end;

procedure TOF_BTJOUCON.OnUpdate;
begin
  inherited;

  ChargeLesConsos;

  CompletelesInfos;

  if TOBBJC <> nil then TFStat(Ecran).LaTOB := TOBBJC;

end;

procedure TOF_BTJOUCON.OnLoad;
Var Aff0 : String;
begin
  inherited;

  DateDeb := StrtoDate(GetControltext('BJCDATE'));
  DateFin := StrtoDate(GetControltext('BJCDATE_'));

  //if TypeAffaire.Value = '' then TypeAffaire.Value := 'A';
  //
  Aff0 := TypeAffaire.Value;
  //
  Aff1 := GetControlText('BJCAFFAIRE1');
  Aff2 := GetControlText('BJCAFFAIRE2');
  Aff3 := GetControlText('BJCAFFAIRE3');
  Aff4 := GetControlText('BJCAVENANT');
  if AFF4 = '' then Aff4 := '0';

  //On recharge l'affaire en une seule zone...
  //if GetControlText('BJCAFFAIRE') = '' then
  //begin
  CodeAffaire := RegroupePartiesAffaire(Aff0, Aff1, Aff2, Aff3, Aff4);
  SetcontrolText('BJCAFFAIRE', CodeAffaire);
  //end;
  //else
  //CodeAffaire := GetControlText('BJCAFFAIRE');

  BTPCodeAffaireDecoupe (CodeAffaire,Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);

  if TypeAffaire.Value = 'A' then
    LibAffaire.Caption := 'Chantier :'
  else if TypeAffaire.Value = 'I' then
    LibAffaire.Caption := 'Contrat :'
  else if TypeAffaire.Value = 'W' then
    LibAffaire.Caption := 'Appel :'
  else
    LibAffaire.Caption := 'Affaire :';

  CodeAffaire1 := Aff1;
  CodeAffaire2 := Aff2;
  CodeAffaire3 := Aff3;

  Ressource := GetControltext('BJCRESSOURCE');

  Nature := GetControltext('BJCNATURE');
  if Ok_FullScreen then
  begin
  	ShowWindow(TForm(Ecran).Handle, SW_SHOWMAXIMIZED);
  	// PassagePleinEcran(THForm(Ecran)); //TForm(Ecran).WindowState := wsMaximized;
  end;


end;

procedure TOF_BTJOUCON.OnArgument(S: string);
var stArgument,Critere,Champ,Valeur : string;
		X : integer;
begin
  inherited;

  LanceAuto := False;
  Ok_fullScreen := false;

  SetControlText('XX_WHERE', '');

  TypeAffaire:= THMultiValComboBox(Getcontrol('AFFAIRE0'));
  TypeAffaire.OnChange := OnChangeTypeAffaire;
  //if TypeAffaire.Value = '' then TypeAffaire.Value := 'A';

  bSelectAff := TToolBarButton97(GetControl('BSELECTAFF1'));
  Bselectaff.OnClick := OnClickSelectAff;

  // traitement des paramètres
  stArgument := S;
  Critere := (Trim (ReadTokenSt (stArgument)));

  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos ('=', Critere) ;
      if x <> 0 then
      begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        ControleChamp(Champ, Valeur);
      end;
    end;
    ControleCritere(Critere);
    Critere := (Trim (ReadTokenSt (stArgument))) ;
  end;

  // Gestion des évènements du TB
  LibAffaire := THLabel(GetControl('TBTB_AFFAIRE'));

  TV1 := TTobViewer(GetControl('TV'));
  TV1.OnDblClick := OnDblClickTV;
  TOBBJC := TOB.create('Ma Tob', nil, -1);

end;

Procedure TOF_BTJOUCON.OnChangeTypeAffaire(Sender : TObject);
Var Aff0    : string;
    Tmp     : string;
    Critere : string;
Begin

  Critere := TypeAffaire.Value;
  Tmp     := Uppercase(Trim(ReadTokenSt(Critere)));

  while Tmp <> '' do
  begin
     Aff0 := Aff0 + Tmp + ',';
     Tmp := uppercase(Trim(ReadTokenSt(Critere)));
  end;

  Aff0 := Copy(Aff0, 1, Length(Aff0)-1);

  if Aff0 <> STypeaff then
  begin
    SetControlText('BJCAFFAIRE1', '');
    SetControlText('BJCAFFAIRE2', '');
    SetControlText('BJCAFFAIRE3', '');
    SetControlText('BJCAVENANT',  '');
  end;

  if Aff0 = 'A' then
  Begin
    LibAffaire.Caption := 'Chantier :';
    SetControlText('BJCAFFAIRE0', 'A');
  end
  else if Aff0 = 'I' then
  Begin
    LibAffaire.Caption := 'Contrat :';
    SetControlText('BJCAFFAIRE0', 'I');
  end
  else if Aff0 = 'W' then
  begin
    LibAffaire.Caption := 'Appel :';
    SetControlText('BJCAFFAIRE0', 'W');
  end
  else
  Begin
    LibAffaire.Caption := 'Affaire :';
    SetControlText('BJCAFFAIRE0', '');
  end;

  STypeaff := Aff0;

end;

Procedure TOF_BTJOUCON.ControleChamp(Champ, Valeur : String);
Var Aff0 : string;
begin

  if Champ = 'AFFAIRE' then
  begin
    SetControlText('BJCAFFAIRE', Valeur);
    BTPCodeAffaireDecoupe (Valeur,Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);
    TypeAffaire.Value := Aff0;
    SetControlText('BJCAFFAIRE1', Aff1);
    SetControlText('BJCAFFAIRE2', Aff2);
    SetControlText('BJCAFFAIRE3', Aff3);
    SetControlText('BJCAVENANT', Aff4);
    //FV1 : 16/12/2016 - FS#2293 - Journal de conso à partir fiche affaire bloqué les éléments de lancement.
    TypeAffaire.Enabled := False;
    SetControlProperty('BJCAFFAIRE1', 'Enabled', False);
    SetControlProperty('BJCAFFAIRE2', 'Enabled', False);
    SetControlProperty('BJCAFFAIRE3', 'Enabled', False);
    SetControlProperty('BJCAVENANT',  'Enabled', False);
    SetControlProperty('BEffaceAFF1', 'Visible', False);
    SetControlProperty('BSelectAff1', 'Visible', False);
  end;

  if Champ = 'DATEDEB' then
  begin
    DateDeb := StrTodateTime(Valeur);
    SetControltext('BJCDATE', Valeur);
  end;

  if Champ = 'DATEFIN' then
  begin
    DateFin := StrToDateTime(Valeur);
    SetControltext('BJCDATE_', Valeur);
  end;

  if Champ = 'NATUREMOUV' then SetControltext('BJCNATURE',Valeur);

end;

Procedure TOF_BTJOUCON.ControleCritere(Critere : String);
begin

  If Critere = 'AVANCE' then SetControlText('XX_WHERE', 'AND BCO_NATUREMOUV IN ("EXT","MO")');

  If Critere = 'FULLSCREEN' then Ok_FullScreen := True

end;

procedure TOF_BTJOUCON.OnClose;
begin
  inherited;
  if TOBBJC <> nil then FreeAndNil(TOBBJC);
end;

//******************************************************************************
//********************* Zooms / double clic  ***********************************
//******************************************************************************

procedure TOF_BTJOUCON.OnDblClickTV(Sender: TObject);
begin
end;

procedure TOF_BTJOUCON.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  Aff0 := THEdit(GetControl('BJCAFFAIRE0'));
  Aff  := THEdit(GetControl('BJCAFFAIRE'));
  Aff1 := THEdit(GetControl('BJCAFFAIRE1'));
  Aff2 := THEdit(GetControl('BJCAFFAIRE2'));
  Aff3 := THEdit(GetControl('BJCAFFAIRE3'));
  Aff4 := THEdit(GetControl('BJCAVENANT'));
end;

procedure TOF_BTJOUCON.ChargeLesConsos;
var Req : string;
    QQ: Tquery;
begin

  Req := TFStat(Ecran).stSQL;
  // On met à jour la TOB uniquement pour les chantiers
  // ayant des mouvements dans les dates sélectionnées
  Req := Req + ' AND BCO_DATEMOUV BETWEEN "' + UsDateTime(DateDeb) +
    '" AND "' + UsDateTime(DateFin) + '"';

  CodeAffaire1 := GetControlText('BJCAFFAIRE1');
  CodeAffaire2 := GetControlText('BJCAFFAIRE2');
  CodeAffaire3 := GetControlText('BJCAFFAIRE3');
              
  //FV1 - 15/12/2016 - FS#2290 - LE SANITAIRE FRANCAIS - journal de consommations ne fonctionne plus
  //if CodeAffaire <> '' then
  //begin
  //  Req := Req + ' AND BCO_AFFAIRE = "' + CodeAffaire + '"';
  //end
  //else
  //begin
  if TypeAffaire.value <> '' then Req := Req + ' AND BCO_AFFAIRE0 IN (' + FormatMultiValComboforSQL(TypeAffaire.Value) + ')';
  //
  if CodeAffaire1 <> '' then Req := Req + ' AND BCO_AFFAIRE1 = "' + CodeAffaire1 + '"';
  if CodeAffaire2 <> '' then Req := Req + ' AND BCO_AFFAIRE2 = "' + CodeAffaire2 + '"';
  if CodeAffaire3 <> '' then Req := Req + ' AND BCO_AFFAIRE3 = "' + CodeAffaire3 + '"';
  //end;

  if Ressource <> '' then
  begin
    Req := Req + ' AND BCO_RESSOURCE = "' + Ressource + '"';
  end;
  if Nature <> '' then
  begin
    Req := Req + ' AND BCO_NATUREMOUV = "' + Nature + '"';
  end;

  QQ := OpenSQL(Req, true,-1,'',true);
  TOBBJC.LoadDetailDB('', '', '', QQ, false);
  ferme(QQ);

end;

procedure TOF_BTJOUCON.AddChampsSupLoc (TOBB : TOB);
begin
	TOBB.AddChampSupValeur ('Fournisseur','');
	TOBB.AddChampSupValeur ('Provenance','');
  //
end;

//FV1 - 29/03/2018 : FS#3025 - ECHAFAUDAGE SERVICE : Permettre d'afficher le libellé des tables libres dans le journal des conso

procedure TOF_BTJOUCON.AddChampsSupLibTableLibre (TOBB : TOB);
begin
  //
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res1', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res2', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res3', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res4', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res5', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res6', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res7', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res8', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_Res9', '');
  TOBB.AddChampSupValeur ('Lib_Table_Libre_ResA', '');
  //
end;

procedure TOF_BTJOUCON.AlimenteInfoProv (TOBL,TOBF : TOB);
var prov : string;
begin
	TOBL.putValue('Fournisseur',TOBF.GetValue('GL_TIERS'));
  Prov := RechDom ('GCNATUREPIECEG',TOBF.GetValue('BCO_NATUREPIECEG'),false)+' '+
  				IntToStr(TOBF.GetValue('BCO_NUMERO'))+ ' du '+DateToStr(TOBF.GetValue('GL_DATEPIECE'));
	TOBL.putValue('Provenance',prov);
end;

//FV1 - 29/03/2018 : FS#3025 - ECHAFAUDAGE SERVICE : Permettre d'afficher le libellé des tables libres dans le journal des conso
procedure TOF_BTJOUCON.AlimenteLibelleTableLibreRessource(TOBL : TOB);
var Ind     : Integer;
    Numero  : string;
    lib     : string;
begin

  for ind := 1 to 10 do
  begin
    if Ind = 10 then
      Numero := 'A'
    else
      Numero := IntToStr(Ind);
    lib    := RechDom ('AFTLIBRERES'+Numero, TOBL.GetValue('TABLE_LIBRE_RESSOURCE'+Numero),false);
	  TOBL.putValue('Lib_Table_Libre_Res'+Numero, Lib);
  end;

end;

procedure TOF_BTJOUCON.CompletelesInfos;
var Indice : integer;
		TOBL,TOBF : TOB;
    TOBProv : TOB;
    QQ : TQuery;
    Req : String;
begin
	TOBProv := TOB.Create ('LES LIGNES FOUR',nil,-1);
  for Indice := 0 to TOBBJC.detail.count -1 do
  begin
    TOBL := TOBBJC.detail[Indice];
    AddChampsSupLibTableLibre(TOBL);
    if (TCheckBox(GetCOntrol('INCLUSINFOFOU')).Checked) then
    begin
      AddChampsSupLoc (TOBL);
      if (TOBL.GetValue('BCO_NATUREMOUV')='FOU') or (TOBL.GetValue('BCO_NATUREMOUV')='EXT') then
      begin
        TOBProv.ClearDetail;
        Req := 'SELECT GL_TIERS,GL_DATEPIECE,BCO_NATUREPIECEG,BCO_SOUCHE,BCO_NUMERO,BCO_INDICEG '+
                       'FROM CONSOMMATIONS LEFT JOIN LIGNE ON GL_NATUREPIECEG=BCO_NATUREPIECEG AND '+
                       'GL_SOUCHE=BCO_SOUCHE AND GL_NUMERO=BCO_NUMERO AND GL_INDICEG=BCO_INDICEG ';
        if (TOBL.GetValue('BCO_NATUREPIECEG')='AF') or (TOBL.GetValue('NATUREPIECEG')='AFS') then     //Pas de lienvente pour les avoirs forunisseurs
          Req := Req + 'WHERE BCO_NUMMOUV='+floatToStr(TOBL.GetValue('BCO_NUMMOUV'))
        else
          Req := Req + 'WHERE BCO_LIENVENTE='+floatToStr(TOBL.GetValue('BCO_NUMMOUV'));

        QQ := OpenSql (req,true,-1,'',true);
        if not QQ.eof then
        begin
          TOBProv.LoadDetailDB ('CONSOMMATIONS','','',QQ,false);
          TOBF := TOBProv.findFirst(['BCO_NATUREPIECEG'],['FF'],true);
          if TOBF = nil then
          begin
            TOBF := TOBProv.findFirst(['BCO_NATUREPIECEG'],['AF'],true);
            if TOBF = nil then
            begin
              TOBF := TOBProv.findFirst(['BCO_NATUREPIECEG'],['AFS'],true);
              if TOBF = nil then
              begin
                TOBF := TOBProv.findFirst(['BCO_NATUREPIECEG'],['BLF'],true);
                if TOBF = nil then
                begin
                  TOBF := TOBProv.findFirst(['BCO_NATUREPIECEG'],['LFR'],true);
                  if TOBF <> nil then
                  begin
                    AlimenteInfoProv (TOBL,TOBF);
                  end;
                end
                else
                begin
                  AlimenteInfoProv (TOBL,TOBF);
                end;
              end
              else
              begin
                AlimenteInfoProv (TOBL,TOBF);
              end;
            end
            else
            begin
              AlimenteInfoProv (TOBL,TOBF);
            end;
          end
          else
          begin
            AlimenteInfoProv (TOBL,TOBF);
          end;
        end;
        ferme (QQ);
      end;
    end;
    //FV1 - 29/03/2018 : FS#3025 - ECHAFAUDAGE SERVICE : Permettre d'afficher le libellé des tables libres dans le journal des conso
    AlimenteLibelleTableLibreRessource(TOBL);
  end;
  if TOBProv <> nil then FreeAndNil(TOBProv);
end;

Procedure TOF_BTJOUCON.OnClickSelectAff(Sender : Tobject);
Var Stchamps        : String;
    StArgument      : String;
    Tmp             : String;
    Aff0            : string;
    Critere         : String;
begin

  StArgument := 'ACTION=CONSULTATION';
  Stargument := Stargument + ';NOCHANGETIERS';
  Stargument := Stargument + ';ACTION=RECH';
  Stargument := Stargument + ';NOFILTRE';
  StArgument := Stargument + ';NOAFFETAT';

  Critere := TypeAffaire.Value;
  Tmp     := Uppercase(Trim(ReadTokenSt(Critere)));

  while Tmp <> '' do
  begin
     Aff0 := Aff0 + Tmp + ',';
     Tmp := uppercase(Trim(ReadTokenSt(Critere)));
  end;

  Aff0 := Copy(Aff0, 1, Length(Aff0)-1);

  if (Aff0 <> '') or
     (Aff1 <> '') or
     (Aff2 <> '') or
     (Aff3 <> '') then
  begin
    if (Aff0 <> '')               then StChamps := StChamps + ';AFF_AFFAIRE0 IN ('' + FormatMultiValComboforSQL(TypeAffaire.Value) + '')';
    if (Aff1 <> '')               then StChamps := StChamps + ';AFF_AFFAIRE1=' + Aff1;
    if (Aff2 <> '')               then StChamps := StChamps + ';AFF_AFFAIRE2=' + Aff2;
    if (Aff3 <> '')               then StChamps := StChamps + ';AFF_AFFAIRE3=' + Aff3;
    if (Aff4 <> '00')             then
    begin
      if (Aff4 <> '') then StChamps := StChamps + ';AFF_AVENANT=' + Aff4;
    end;
  end;

  if Stchamps <> '' then Stchamps := Stchamps + ';';

  if Aff0 = 'A' then
  begin
     StChamps := Stchamps + 'AFF_STATUTAFFAIRE=AFF';
     StArgument := Stargument + ';NOCHANGESTATUT;STATUT=AFF';
   	 tmp := AGLLanceFiche('BTP', 'BTAFFAIRE_MUL', StChamps, '', StArgument);
  end
  else if Aff0 = 'I' then
  begin
    StChamps := Stchamps + 'AFF_STATUTAFFAIRE=INT';
    StArgument := Stargument + ';NOCHANGESTATUT;STATUT=INT';
    tmp := AGLLanceFiche('BTP', 'BTCONTRAT_MUL', StChamps, '', StArgument);
  end
  else if Aff0 = 'W' then
  begin
    StChamps := Stchamps + 'AFF_STATUTAFFAIRE=APP';
    StArgument := Stargument + ';NOCHANGESTATUT;STATUT=APP;ETAT=CL1,FAC,FIN,REA';
    tmp := AGLLanceFiche('BTP', 'BTMULAPPELS', StChamps, '', StArgument);
  end
  else
  begin
    tmp := AGLLanceFiche('BTP', 'BTAFFAIRE_MUL', StChamps, '', 'STATUT=AFF' + StArgument);
  end;

  if tmp <> '' then
  begin
    CodeAffaire := ReadTokenSt(tmp);
    BTPCodeAffaireDecoupe(Codeaffaire,Aff0,Aff1,Aff2,Aff3,Aff4,taCreat, False);
    SetControlText('BJCAFFAIRE0', Aff0);
    TypeAffaire.Value := Aff0 + ';';
    SetControlText('BJCAFFAIRE',  CodeAffaire);
    SetControlText('BJCAFFAIRE1', Aff1);
    SetControlText('BJCAFFAIRE2', Aff2);
    SetControlText('BJCAFFAIRE3', Aff3);
    SetControlText('BJCAVENANT',  Aff4);
  end;


end;

initialization
  registerclasses([TOF_BTJOUCON]);
end.
