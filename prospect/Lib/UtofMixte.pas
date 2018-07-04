{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 27/09/2006
Modifié le ... : 27/09/2006
Description .. : Ancêtre de uTof
Suite ........ : permettre la gestion sur les muls mixte (tiers - suspects) des
Suite ........ : champs informations complémentaires en critères
Mots clefs ... : TOF; CRITERE; MIXTE; INFOS COMPL
*****************************************************************}
Unit UtofMixte ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     utilSelection,   //MulCreerOnglet
     ExtCtrls,    //Tbevel
     HEnt1,
     EntRT,     //La variable VH_RT
     HTB97,
     UtilGC,          // GCMAJChampLibre
     KPMGutil,
     Variants,
     HMsgBox,
{$ifdef AFFAIRE}
    UTOFAFTRADUCCHAMPLIBRE,     //TOF_AFTRADUCCHAMPLIBRE
{$endif}
     UTOF ;


//les paramètres
  Type Param_Mixte = record
    Suffixe             : String;       //suffixes des champs
    Action              : variant;      //mode d'ouverture des fiches
    TypeFiche           : String;       //Ciblage / recherche / ...
  end;

Type
{$ifdef AFFAIRE}
                //PL le 18/05/07 pour faire affectation depuis ressource si paramétré
     TOF_MIXTE = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_MIXTE = Class (TOF)
{$endif}
   private
    TobCorSuspect       : Tob;
    ExisteVal           : boolean;
    ExisteTexte         : boolean;
    ExisteBool          : boolean;
    ExisteDate          : boolean;
    ExisteTable         : boolean;
    ExisteTiersTable    : Boolean;
    IdxTiersTable       : integer;
    IdxTable            : integer;
    IdxTexte            : integer;
    BZoom               : TToolbarButton97;

   public
    ParamMixte          : Param_Mixte;        // variable pour les paramètres

    procedure GereChLibLibre ( TypeCh : String ; Index : integer );
    Function GereChSqlLibre ( TypeCh : String; Index : integer) : String;
    procedure UpdateWhere (LeWhere : String);

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    //les événements
    procedure BZoomOnClick (Sender : Tobject);

  end ;


Implementation

uses uTOFComm, utilrt, DateUtils, TntStdCtrls;

procedure TOF_MIXTE.OnArgument (S : String ) ;
Var
  I                     : integer;
  Tabs                  : TTabSheet;
  strSql                : String;
  XX_WHereOld,XX_WHereNew,XX_whereComplet : string;
  posAND                : integer;
  bWhereDifferents      : boolean;

begin
//  Inherited ;

  ExisteVal             := False;
  ExisteTexte           := False;
  ExisteBool            := False;
  ExisteDate            := False;
  ExisteTable           := False;
  ExisteTiersTable      := False;
  IdxTiersTable         := 0;
  IdxTable              := 0;
  IdxTexte              := 0;

  BZoom                 := TToolbarButton97(GetControl('BZOOM'));
  If Assigned(BZoom) then
  begin
    BZoom.Visible       := True;
    BZoom.OnClick       := BZoomOnClick;
  end;

  //chargement de la table des correspondances
  StrSql                := 'SELECT * FROM PARSUSPECTCOR';
  TobCorSuspect         := Tob.Create('Les correspondances', nil, -1);
  TobCorSuspect.LoadDetailFromSQL(StrSql);

  //TJA 06/07/2007   restriction des fiches Tiers
  //PL 13/07/07
  //Fq10446
  bWhereDifferents := false;
  XX_WhereNew := '';
  XX_whereComplet := RTXXWhereConfident('CON',true);
  XX_WHereOld := VH_RT.RTConfWhereConsult;
  if (length (XX_whereComplet) <> length (XX_WHereOld)) then
    bWhereDifferents := true;

  if (XX_WhereOld <> '') and (XX_whereComplet <> '') then
    begin
      XX_WhereNew := XX_whereComplet;
      if bWhereDifferents then
      begin
        posAND := pos ('AND', XX_WhereOld);
        if (posAND <> 0) and (posAND < 4) then // on supprime le AND au début
          XX_WhereOld := copy(XX_WhereOld, posAND + 3, length(XX_WhereOld));
        XX_WhereNew := XX_WhereOld;
        XX_WhereNew := ' (' + XX_WhereNew + ' OR ' + ParamMixte.Suffixe + '_NATURE = "SUS")';
        XX_whereComplet := StringReplace(XX_whereComplet, XX_WhereOld, XX_WhereNew, [rfReplaceAll]);
      end
      else
      begin
        XX_whereComplet := XX_whereComplet + ' OR ' + ParamMixte.Suffixe + '_NATURE = "SUS"';
      end;

    XX_whereComplet := StringReplace(XX_whereComplet, 'T_TIERS', ParamMixte.Suffixe + '_CODE', [rfReplaceAll, rfIgnoreCase]);
    XX_whereComplet := StringReplace(XX_whereComplet, ' T_', ' ' + ParamMixte.Suffixe + '_', [rfReplaceAll, rfIgnoreCase]);
    XX_whereComplet := StringReplace(XX_whereComplet, '(T_', '(' + ParamMixte.Suffixe + '_', [rfReplaceAll, rfIgnoreCase]);

    SetControlText('XX_WHERE', XX_whereComplet);
    end
  else if (XX_whereComplet <> '') then    //mcd 15/07/08 15411 si dossier sans restriciton, pas OK
     begin
       XX_whereComplet := StringReplace(XX_whereComplet, 'T_TIERS', ParamMixte.Suffixe + '_CODE', [rfReplaceAll, rfIgnoreCase]);
       SetControlText('XX_WHERE', XX_whereComplet);
     end;

  //attribution des libellés
  for I := 0 to 19 do
  begin
    GereChLibLibre('VAL', I);
    GereChLibLibre('DATE', I);
    GereChLibLibre('BOOL', I);
  end;

  for I := 0 to 29 do
    GereChLibLibre('TEXTE', I);
    
  for I := 0 to 35 do
    GereChLibLibre('TABLE', I);

  TFMul(Ecran).Pages.ActivePage := Nil;

  inherited;                // héritage à ce moment pour avoir la gestion des tablettes en correspondances
  //on ferme les onglets s'il n'y a pas de champs en correspondance

  Tabs                  := TTabSheet(GetControl('PVALEURLIBRE'));
  Tabs.TabVisible       := ExisteVal;
  Tabs                  := TTabSheet(GetControl('PTEXTELIBRE'));
  Tabs.TabVisible       := ExisteTexte;
  Tabs                  := TTabSheet(GetControl('PBOOLLIBRE'));
  Tabs.TabVisible       := ExisteBool;
  Tabs                  := TTabSheet(GetControl('PDATELIBRE'));
  Tabs.TabVisible       := ExisteDate;
  Tabs                  := TTabSheet(GetControl('PTABLELIBRE'));
  Tabs.TabVisible       := ExisteTable;
  Tabs                  := TTabSheet(GetControl('PTABLETIERS'));
  Tabs.TabVisible       := ExisteTiersTable;

 //PL le 18/05/07 : bloc lié à la GI
  if (GetControl('YTC_RESSOURCE1') <> nil)  then
    begin
    if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
    else begin
      GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
      if not (ctxscot in V_PGI.PGICOntexte) then
        begin
        SetControlVisible ('T_MOISCLOTURE',false);
        SetControlVisible ('T_MOISCLOTURE_',false);
        SetControlVisible ('TT_MOISCLOTURE',false);
        SetControlVisible ('TT_MOISCLOTURE_',false);
        end;
      end;
    end;

  {$IFDEF GCGC}
  //FQ10467 gestion des commerciaux // PL le 18/05/07 : on affine le test
  if not (ctxaffaire in V_PGI.PGICONTEXTE) and not GereCommercial then
  begin
    if assigned(GetControl(ParamMixte.Suffixe + '_REPRESENTANT')) then
      begin
        SetControlVisible(ParamMixte.Suffixe  + '_REPRESENTANT', False);
        SetControlVisible('T' + ParamMixte.Suffixe  + '_REPRESENTANT', False);
      end;
  end;
  {$ENDIF GCGC}


end ;

procedure TOF_MIXTE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MIXTE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MIXTE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MIXTE.OnLoad ;
Var
  I                     : integer;
  StrWhere              : String;
  StrW                  : String;

begin
  Inherited ;
  if Not Assigned(GetControl('XX_WHERE')) then
    Exit;

  // recherche des champs renseignés pour les inclure dans le Where
  For i := 0 to 19 do
  begin
    StrW                := '';
    StrW                := GereChSqlLibre('VAL', I);
    if (StrW) <> '' then
    begin
      if StrWhere <> '' then
        StrWhere        := StrWhere + ' AND ' + StrW
      else
        StrWhere        := StrW;
    end;

    StrW                := '';
    StrW                := GereChSqlLibre('BOOL', I);
    if (StrW) <> '' then
    begin
      if StrWhere <> '' then
        StrWhere        := StrWhere + ' AND ' + StrW
      else
        StrWhere        := StrW;
    end;

    StrW                := '';
    StrW                := GereChSqlLibre('DATE', I);
    if (StrW) <> '' then
    begin
      if StrWhere <> '' then
        StrWhere        := StrWhere + ' AND ' + StrW
      else
        StrWhere        := StrW;
    end;

  end;

  for I := 0 to 29 do
  begin
    StrW                := '';
    StrW                := GereChSqlLibre('TEXTE', I);
    if (StrW) <> '' then
    begin
      if StrWhere <> '' then
        StrWhere        := StrWhere + ' AND ' + StrW
      else
        StrWhere        := StrW;
    end;
  end;

  for I := 0 to 35 do
  begin
    StrW                := '';
    StrW                := GereChSqlLibre('TABLE', I);
    if (StrW) <> '' then
    begin
      if StrWhere <> '' then
        StrWhere        := StrWhere + ' AND ' + StrW
      else
        StrWhere        := StrW;
    end;
  end;

//  SetControlText('XX_WHERE', StrWhere);
  if Assigned(GetControl('XX_WHERESPE')) then
    SetControlText('XX_WHERESPE', Strwhere);
end ;


procedure TOF_MIXTE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MIXTE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MIXTE.OnCancel () ;
begin
  Inherited ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. : Affectation du libellé sur le champ
Mots clefs ... :
*****************************************************************}
procedure TOF_MIXTE.GereChLibLibre(TypeCh: String; Index: integer);
var
  Lechamp               : String;
  TobRes                : Tob;
  LaCombo               : THValComboBox;
  LOngletSup            : TTabSheet;
  LeLabel               : THLabel;
  LEdit                 : THEdit;
  StrChamp              : String;

begin

  TypeCh                := UpperCase_(TypeCh);
  Lechamp               := 'T' + TypeCh + IntToStr(Index);
  if typeCh = 'BOOL' then
    Lechamp             := TypeCh + IntToStr(Index);

//  if pos('_6', LeChamp) > 0 then
//    LeChamp             := StringReplace(LeChamp, '_6', '_', [rfReplaceAll]);

  TobRes:=TobCorSuspect.FindFirst(['RSP_CHSUSPECT'],['RSC_RSCLIB'+TypeCh+IntToStr(Index)],False);
  if (TobRes = nil) and (TypeCh = 'TABLE') then     // peut être 6RSCLIB
    TobRes              := TobCorSuspect.FindFirst(['RSP_CHSUSPECT'],['RSC_6RSCLIBTABLE'+IntToStr(Index)],False);

  if Tobres<>nil then
  begin

    //on affiche le libellé du champ tiers associé
    if (TypeCh <> 'BOOL') and (TypeCh <> 'TABLE') and (TypeCh <> 'TEXTE') then
      SetControlText(Lechamp, TobRes.GetValue('RSP_LIBELLE'));
    if (TypeCh = 'BOOL') then
      SetControlCaption(Lechamp, TobRes.GetValue('RSP_LIBELLE'));

    if TypeCh = 'VAL' then
      ExisteVal         := True;

    if TypeCh = 'TEXTE' then
    begin
      if IdxTexte >= 12 then
      begin
        if IdxTexte < 24 then
        begin
          LOngletSup      := TTabSheet(GetControl('PTEXTELIBRE2'));
          if not Assigned(LOngletSup) then
          begin
            MulCreerOnglet(True, Tfmul(Ecran), Ttabsheet(Getcontrol('PTEXTELIBRE')).PageIndex + 1, LOngletSup, false);
            with LOngletSup do
            begin
              Name        := 'PTEXTELIBRE2';
              Caption     := 'Textes libres #2';
              PageIndex   := TTabSheet(GetControl('PTEXTELIBRE')).PageIndex + 1;
            end;
            StrChamp      := 'BEVELOl' + IntToStr(Ttabsheet(Getcontrol('PTEXTELIBRE')).PageIndex + 1);
            SetControlProperty(StrChamp, 'NAME', 'BEVELTTL2');
          end;
        end
        else
        begin
          LOngletSup      := TTabSheet(GetControl('PTEXTELIBRE3'));
          if not Assigned(LOngletSup) then
          begin
            MulCreerOnglet(True, Tfmul(Ecran), Ttabsheet(Getcontrol('PTEXTELIBRE')).PageIndex + 1, LOngletSup, false);
            with LOngletSup do
            begin
              Name        := 'PTEXTELIBRE3';
              Caption     := 'Textes libres #3';
              PageIndex   := TTabSheet(GetControl('PTEXTELIBRE2')).PageIndex + 1;
            end;
            StrChamp      := 'BEVELOl' + IntToStr(Ttabsheet(Getcontrol('PTEXTELIBRE')).PageIndex + 1);
            SetControlProperty(StrChamp, 'NAME', 'BEVELTTL3');
          end;
        end;
        LeLabel         := THLabel.Create(Ecran);
        With LeLabel do
        begin
          Name          := 'TTEXTE' + IntToStr(IdxTexte);
          Parent        := LOngletSup;
          Caption       := 'Texte libre ' + IntToStr(IdxTexte);
          StrChamp      := 'TTEXTE' + IntToStr(IdxTexte - 12);
          Left          := Getcontrol(StrChamp).Left;
          Top           := Getcontrol(StrChamp).Top;
          Width         := Getcontrol(StrChamp).Width;
          Height        := GetControl(StrChamp).Height;
        end;
        LEdit           := THEdit.Create(Ecran);
        With LEdit do
        begin
          Name          := 'TEXTE' + IntToStr(IdxTexte);
          Parent        := LOngletSup;
          Text          := '';
          StrChamp      := 'TEXTE' + IntToStr(IdxTexte - 12);
          Left          := Getcontrol(StrChamp).Left;
          Top           := Getcontrol(StrChamp).Top;
          Width         := Getcontrol(StrChamp).Width;
          Height        := GetControl(StrChamp).Height;
        end;
      end;

      ExisteTexte       := True;
      Lechamp           := 'T' + TypeCh + IntToStr(IdxTexte);
      LEdit             := THEdit(GetControl(TypeCh + IntToStr(IdxTexte)));
      if Assigned(LEdit) then
      begin
        SetControlText('T' + TypeCh + IntToStr(IdxTexte), TobRes.GetValue('RSP_LIBELLE'));
        TobRes.AddChampSupValeur('LECHAMPTASSE', TypeCh + IntToStr(IdxTexte));
        LEdit.Visible   := True;
        Inc(IdxTexte);
      end;
    end;

    if TypeCh = 'BOOL' then
      ExisteBool        := True;
    if TypeCh = 'DATE' then
      ExisteDate        := True;
    if TypeCh = 'TABLE' then
    begin
      // TABLE TIERS
      if TobRes.GetValue('RSP_CHTIERS') <> '' then
      begin
//        if idxTable = 10 then
//          exit;
        ExisteTiersTable  := True;
        LeChamp           := 'TTIERS' + TypeCh + IntToStr(IdxTiersTable);
        LaCombo           := THValComboBox(GetControl('TIERS' + TypeCh + IntToStr(IdxTiersTable)));
        if Assigned(LaCombo) then
        begin
          SetControlText('TTIERS' + TypeCh + IntToStr(IdxTiersTable), TobRes.GetValue('RSP_LIBELLE'));
          TobRes.AddChampSupValeur('LECHAMPTASSE', 'TIERS' + TypeCh + IntToStr(IdxTiersTable));
          LaCombo.DataType  := RenvoiTabletteCor(TobCorSuspect, TobRes.GetValue('RSP_CHSUSPECT'));
          LaCombo.Visible   := True;
          Inc(IdxTiersTable);
        end;
      end
      else
        //TABLE PROSPECTS
      begin
        // si on doit gérer + de 12 combo pour les prospects
        if (IdxTable >= 12) then
        begin
          if (IdxTable < 24) then
          begin
            LOngletSup  := TTabSheet(GetControl('PTABLELIBRE2'));
            if not assigned(LOngletSup) then
            begin
              MulCreerOnglet(True, Tfmul(Ecran), Ttabsheet(Getcontrol('PTABLELIBRE')).PageIndex + 1, LOngletSup, false);
              with LOngletSup do
              begin
                Name    := 'PTABLELIBRE2';
                Caption := 'Tables libres #2';
                PageIndex := TTabSheet(GetControl('PTABLELIBRE')).PageIndex + 1;
              end;
              StrChamp  := 'BEVELOl' + IntToStr(Ttabsheet(Getcontrol('PTABLELIBRE')).PageIndex + 1);
              SetControlProperty(StrChamp, 'NAME', 'BEVELTL2');
            end;
          end
          else
          begin
            LOngletSup  := TTabSheet(GetControl('PTABLELIBRE3'));
            if not assigned(LOngletSup) then
            begin
              MulCreerOnglet(True, Tfmul(Ecran), Ttabsheet(Getcontrol('PTABLELIBRE')).PageIndex + 1, LOngletSup, false);
              with LOngletSup do
              begin
                Name    := 'PTABLELIBRE3';
                Caption := 'Tables libres #3';
                PageIndex := TTabSheet(GetControl('PTABLELIBRE2')).PageIndex + 1;
                StrChamp  := 'BEVELOl' + IntToStr(Ttabsheet(Getcontrol('PTABLELIBRE')).PageIndex + 1);
                SetControlProperty(StrChamp, 'NAME', 'BEVELTL3');
              end;
            end;
          end;

          LeLabel       := THLabel.Create(Ecran);
          With LeLabel do
          begin
            Name        := 'TTABLE' + IntToStr(IdxTable);
            Parent      := LOngletSup;
            Caption     := 'Table libre ' + IntToStr(IdxTable);
            StrChamp    := 'TTABLE' + IntToStr(IdxTable - 12);
            Left        := Getcontrol(StrChamp).Left;
            Top         := Getcontrol(StrChamp).Top;
            Width       := Getcontrol(StrChamp).Width;
            Height      := GetControl(StrChamp).Height;
          end;
          LaCombo       := THValComboBox.Create(Ecran);
          With LaCombo do
          begin
            Name        := 'TABLE' + IntToStr(IdxTable);
            Parent      := LOngletSup;
            StrChamp    := 'TABLE' + IntToStr(IdxTable - 12);
            Left        := Getcontrol(StrChamp).Left;
            Top         := Getcontrol(StrChamp).Top;
            Width       := Getcontrol(StrChamp).Width;
            Height      := GetControl(StrChamp).Height;
            Vide        := True;
            VideString  := '<<Tous>>';
          end;
        end;


        ExisteTable     := True;
        Lechamp         := 'T' + TypeCh + IntToStr(IdxTable);
        LaCombo         := THValComboBox(GetControl(TypeCh + IntToStr(IdxTable)));
        if Assigned(LaCombo) then
        begin
          SetControlText('T' + TypeCh + IntToStr(IdxTable), TobRes.GetValue('RSP_LIBELLE'));
          TobRes.AddChampSupValeur('LECHAMPTASSE', TypeCh + IntToStr(IdxTable));
          LaCombo.DataType  := RenvoiTabletteCor(TobCorSuspect, TobRes.GetValue('RSP_CHSUSPECT'));
          LaCombo.Visible   := True;
          Inc(IdxTable);
        end;
      end;

    end;

//    if typeCh <> 'BOOL' then
//      SetControlVisible('T'+TypeCh+IntToStr(Index), True);

    SetControlVisible(Lechamp, True);   //le libellé
    if (TypeCh <> 'TABLE') and (TypeCh <> 'TEXTE') then
      SetcontrolVisible(TypeCh + IntToStr(Index), True);  //le champ data

//    if not (TypeCh = 'TEXTE') or not (TypeCh = 'TABLE') or not (TypeCh = 'BOOL') then
    if (TypeCh = 'DATE') or (TypeCh = 'VAL') then
    begin
      SetControlVisible('T'+TypeCh+IntToStr(Index)+'_', True);
      SetControlVisible(TypeCh+IntToStr(Index)+'_', True);
    end;


  end
  else  //si pas de correspondance on efface le critère
  begin
    if typeCh <> 'BOOL' then
      SetControlVisible('T'+TypeCh+IntToStr(Index), False);
    if TypeCh = 'TABLE' then
    begin
      SetControlVisible('TIERS' + TypeCh + IntToStr(Index), False);
      SetControlVisible('TTIERS' + TypeCh + IntToStr(Index), False);
      SetControlVisible(TypeCh + IntToStr(Index), False);
      SetControlVisible('T' + TypeCh + IntToStr(Index), False);
    end
    else
      SetControlVisible(TypeCh+IntToStr(Index), False);
    if not (TypeCh = 'TEXTE') or not (TypeCh = 'TABLE') or not (TypeCh = 'BOOL') then
    begin
      SetControlVisible('T'+TypeCh+IntToStr(Index)+'_', False);
      SetControlVisible(TypeCh+IntToStr(Index)+'_', False);
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. : création de l'ordre Sql pour le Where
Mots clefs ... :
*****************************************************************}
function TOF_MIXTE.GereChSqlLibre(TypeCh: String; Index: integer): String;
var
  StrSql                : String;
  TobRes                : Tob;
  Borne1                : String;
  Borne2                : String;
  StrEgalite            : String;

begin
  TypeCh                := UpperCase_(TypeCh);
  StrSql                := '';
  TobRes:=TobCorSuspect.FindFirst(['RSP_CHSUSPECT'],['RSC_RSCLIB'+TypeCh+IntToStr(Index)],False);
  if (TobRes = nil) and (TypeCh = 'TABLE') then     // peut être 6RSCLIB
    TobRes              := TobCorSuspect.FindFirst(['RSP_CHSUSPECT'],['RSC_6RSCLIBTABLE'+IntToStr(Index)],False);
  if TobRes <> nil then
  begin
//    if not Assigned(GetControl(TypeCh + IntToStr(Index))) then
//      exit;

    if TypeCh = 'VAL' then
    begin
      if GetControlText(TypeCh+IntToStr(Index))='' then
        exit;
      Borne1            := GetControlText(TypeCh+IntToStr(Index));
      Borne2            := GetControlText(TypeCh+IntToStr(Index)+'_');
      StrEgalite        := '>=';
    end;
    if TypeCh = 'DATE' then
    begin
      if StrToDate(GetControlText(TypeCh+IntToStr(Index))) = iDate1900 then
        exit;
      Borne1            := '"'+GetControlText(TypeCh+IntToStr(Index))+'"';
      Borne2            := '"'+GetControlText(TypeCh+IntToStr(Index)+'_')+'"';
      StrEgalite        := '>=';
    end;
    if TypeCh = 'TEXTE' then
    begin
//      if GetControlText(TypeCh+IntToStr(Index)) = '' then
      if (TobRes.GetValue('LECHAMPTASSE') = '') or (GetControlText(TobRes.GetValue('LECHAMPTASSE')) = '') then
        exit;
      Borne1            := '"'+ GetControlText(TobRes.GetValue('LECHAMPTASSE')) +'"';
      StrEgalite        := '=';
    end;

    if TypeCh = 'TABLE' then
    begin
      if (TobRes.GetValue('LECHAMPTASSE') = '') or (GetControlText(TobRes.GetValue('LECHAMPTASSE')) = '') then
        exit;
      Borne1            := '"'+GetControlText(TobRes.GetValue('LECHAMPTASSE'))+'"';
      StrEgalite        := '=';
    end;

    if TypeCh = 'BOOL' then
    begin
      if GetCheckBoxState(TypeCh+IntToStr(Index)) = CbUnchecked then
        exit;
      Borne1            := '"X"';
      StrEgalite        := '=';
    end;

    //Sql suspect
    StrSql              := '(('+TobRes.getvalue('RSP_CHSUSPECT')+StrEgalite+Borne1;
    if (TypeCh = 'VAL') or (TypeCh = 'DATE') then
    begin
      StrSql            := StrSql + ' AND '+TobRes.getvalue('RSP_CHSUSPECT')+'<='+Borne2;
    end;
    StrSql := StrSql + ') OR (';
    //Sql tiers
    if TobRes.GetValue('RSP_CHTIERS')<>'' then
      StrSql            := StrSql +TobRes.getvalue('RSP_CHTIERS')+StrEgalite+Borne1
    else
      StrSql            := StrSql +TobRes.getvalue('RSP_CHCOMPL')+StrEgalite+Borne1;
    if (TypeCh = 'VAL') or (TypeCh = 'DATE') then
    begin
      if TobRes.GetValue('RSP_CHTIERS')<>'' then
        StrSql          := StrSql + ' AND '+TobRes.getvalue('RSP_CHTIERS')+'<='+Borne2
      else
        StrSql          := StrSql + ' AND '+TobRes.getvalue('RSP_CHCOMPL')+'<='+Borne2;
    end;
    StrSql              := StrSql + '))';
  end;

  Result                := StrSql;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 27/09/2006
Modifié le ... :   /  /
Description .. : met à jour le 'XX_WHERE' en prenant compte des critères
Suite ........ : informations complémentaires
Mots clefs ... :
*****************************************************************}
procedure TOF_MIXTE.UpdateWhere(LeWhere: String);
var
  StrW                  : String;
  StrWhere              : String;

begin
  if not Assigned(GetControl('XX_WHERESPE')) then
    Exit;

  StrW                  := GetControlText('XX_WHERESPE');
  if StrW <> '' then
    StrWhere            := LeWhere + ' AND ' + StrW
  else
    StrWhere            := LeWhere;

  SetControlText('XX_WHERE', StrWhere);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 14/06/2007
Modifié le ... :   /  /
Description .. : Bouton de recherche sue le lien
Mots clefs ... : SUSPECT ; TIERS
*****************************************************************}
procedure TOF_MIXTE.BZoomOnClick(Sender: Tobject);
var
  StrCode               : String;
  StrNature             : String;
  StrAuxiliaire         : String;

begin

  if ParamMixte.Suffixe  = '' then
    exit;

  StrCode               := ParamMixte.Suffixe  + '_CODE';
  if VarIsNull(GetField(StrCode)) then
    exit;

  Strauxiliaire         := ParamMixte.Suffixe  + '_AUXILIAIRE';

  StrNature             := ParamMixte.Suffixe  + '_NATURE';
  if GetField(StrNature) = 'SUS' then     // Si fiche Suspect
    V_PGI.DispatchTT(46, ParamMixte.Action, GetField(StrCode), '', 'MONOFICHE')
  else if GetField(StrNature)='CLI' then    // si fiche client
    V_PGI.DispatchTT(8, ParamMixte.Action, GetField(StrAuxiliaire), '', 'MONOFICHE;T_NATUREAUXI=CLI')
  else if GetField(StrNature)='PRO' then    // si fiche prospect
    V_PGI.DispatchTT(8, ParamMixte.Action, GetField(StrAuxiliaire), '', 'MONOFICHE;T_NATUREAUXI=PRO');

end;

Initialization
  registerclasses ( [ TOF_MIXTE ] ) ;
end.
