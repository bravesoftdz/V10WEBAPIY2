{***********UNITE*************************************************
Auteur  ...... : G.Harlez
Créé le ...... : 20/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PARAM_GRPTRAVAIL ()
Mots clefs ... : TOF;PARAM_GRPTRAVAIL
*****************************************************************}
Unit PARAM_GRPTRAVAIL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, 
{$else}
     eMul, 
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTob,
     UTOF,
     HTB97,
     ExtCtrls,
     VIERGE;
Type
  TOF_PARAM_GRPTRAVAIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Obj_LstBxDispo : TListBox;
    Obj_LstBxAff   : TListBox;
    Obj_LstBxTri   : TListBox;
    Obj_BtUP       : TToolbarButton97;
    Obj_BtDOWN     : TToolbarButton97;
    Obj_BtLEFT     : TToolbarButton97;
    Obj_BtRIGHT    : TToolbarButton97;
    Obj_BtRIGHTtri : TToolbarButton97;
    Obj_BtLEFTtri  : TToolbarButton97;
    Obj_RdGrp      : TRadioGroup;
    Obj_RdGrpOrd   : TRadioGroup;
    Obj_EdtTxt     : TEdit;
    Obj_GrpBxopt   : TGroupBox;

    Tb_dispo,Tb_affiche,Tb_tri,Tb_param,Tb_varIN : TOB;

    procedure DoClickButton     (sender : Tobject);
    procedure DoClickRdGrp      (sender : Tobject);
    procedure DoClickOnLstBxTri (sender : Tobject);
    procedure Initialisation    (s_param : string);
    procedure AffichageEcran    (s_mode : String);
  end ;

Implementation
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.OnArgument (S : String ) ;
begin
  Inherited ;
  Obj_LstBxDispo := TListBox(GetControl('LSTBX_COLDISPO'));
  Obj_LstBxAff   := TListBox(GetControl('LSTBX_COLAFF'));

  Obj_LstBxTri   := TListBox(GetControl('LSTBX_COLTRI'));
  if Obj_LstBxTri <> nil then
    Obj_LstBxTri.OnClick := DoClickOnLstBxTri;

  Obj_BtUP       := TToolbarButton97(GetControl('BUP'));
  if Obj_BtUP <> nil then
    Obj_BtUP.OnClick := DoClickButton;

  Obj_BtDOWN     := TToolbarButton97(GetControl('BDOWN'));
  if Obj_BtDOWN <> nil then
    Obj_BtDOWN.OnClick := DoClickButton;

  Obj_BtLEFT     := TToolbarButton97(GetControl('BLEFT'));
  if Obj_BtLEFT <> nil then
    Obj_BtLEFT.OnClick := DoClickButton;

  Obj_BtRIGHT    := TToolbarButton97(GetControl('BRIGHT'));
  if Obj_BtRIGHT <> nil then
    Obj_BtRIGHT.OnClick := DoClickButton;

  Obj_BtRIGHTtri := TToolbarButton97(GetControl('BRIGHT_TRI'));
  if Obj_BtRIGHTtri <> nil then
    Obj_BtRIGHTtri.OnClick := DoClickButton;

  Obj_BtLEFTtri  := TToolbarButton97(GetControl('BLEFT_TRI'));
  if Obj_BtLEFTtri <> nil then
    Obj_BtLEFTtri.OnClick := DoClickButton;

  Obj_RdGrp      := TRadioGroup(GetControl('RDGRP_VISUCHX'));
  if Obj_RdGrp <> nil then
    Obj_RdGrp.OnClick := DoClickRdGrp;

  Obj_RdGrpOrd   := TRadioGroup(GetControl('RDGRP_ORDRE'));
  if Obj_RdGrpOrd <> nil then
    Obj_RdGrpOrd.OnClick := DoClickRdGrp;

  Obj_GrpBxopt := TGroupBox(GetControl('RD_OPT'));

  Obj_EdtTxt := TEdit(GetControl('OPT_LIB'));

  Initialisation(s);
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.OnClose ;
begin
  Inherited ;
  // Libération
  Tb_dispo.Free;
  Tb_affiche.Free;
  Tb_tri.Free;
  Tb_param.Free;
  Tb_varIN.Free;
  Obj_LstBxDispo.Free;
  Obj_LstBxAff.Free;
  Obj_LstBxTri.Free;
  Obj_BtUP.Free;
  Obj_BtDOWN.Free;
  Obj_BtLEFT.Free;
  Obj_BtRIGHT.Free;
  Obj_BtLEFTtri.Free;
  Obj_BtRIGHTtri.Free;
  Obj_RdGrp.Free;
  Obj_RdGrpOrd.Free;
  Obj_EdtTxt.Free;
  Obj_GrpBxopt.Free;
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.Initialisation(s_param : string);
var
  S_sql,s_caption,s_col,s_select,s_orderby,s_ordre : String;
  idx : integer;
  Q : TQuery;
  tb_tmp,tb_temp: tob;
begin
  Tb_varIN   := Tob.create('param entrée',nil,-1);
  Tb_dispo   := Tob.Create('champ dispo',nil,-1);
  Tb_affiche := Tob.Create('champ aff',nil,-1);
  Tb_tri     := Tob.Create('champ tri',nil,-1);
  Tb_param   := Tob.Create('PARAMDONNEES',nil,-1);

  // Apppel de la fiche à partir de l'unite GRPTRAVAIL_TOF.
  // Récuperation des paramétres d'entrées.
  // PARAM_0 : Table
  // PARAM_1 : groupe de données
  // PARAM_2 : Zone de texte libre pour le libellé du radio bouton

  idx := 0;
  while s_param <> '' do
  begin
    s_col := READTOKENPipe(s_param,'|');
    tb_tmp := tob.create('',Tb_varIN,-1);
    tb_tmp.AddChampSupValeur('PARAM_'+inttostr(idx),s_col);
    inc(idx);
  end;

  // Gestion de la liste "Champ disponible"
  S_sql := 'select dh_nomchamp,dh_libelle '+
           'from dechamps,detables '+
           'where dt_nomtable = "'+Tb_varIN.Detail[0].GetString('PARAM_0')+'" '+
           'and dt_prefixe = dh_prefixe';
  Tb_dispo.LoadDetailFromSQL(S_sql);

  // Gestion de la liste "Champ à afficher". récupération du paramétrage existant.
  s_sql := 'select * '+
           'from paramdonnees '+
           'where pmd_grpnom = "'+Tb_varIN.Detail[1].GetString('PARAM_1')+'" '+
           'and pmd_nomtable = "'+Tb_varIN.Detail[0].GetString('PARAM_0')+'"';
  Q := OpenSql(s_sql,FALSE);
  Tb_param.LoadDetailDB('PARAMDONNEES','','',Q,FALSE,FALSE);
  Ferme(Q);

  // Mise en place des colonnes en fonction du paramétrage.
  for idx:=0 to Tb_param.Detail.Count -1 do
  begin
    s_select := Tb_param.Detail[idx].GetString('PMD_SELECT');
    while s_select <> '' do
    begin
      s_col := READTOKENPipe(s_select,',');

      tb_tmp := Tb_dispo.FindFirst(['DH_NOMCHAMP'],[s_col],FALSE);
      if tb_tmp <> nil then
      begin
        Obj_LstBxAff.Items.Add(tb_tmp.GetString('DH_LIBELLE'));
        tb_temp := tob.create('',Tb_affiche,-1);
        tb_temp.AddChampSupValeur('DH_NOMCHAMP',tb_tmp.GetString('DH_NOMCHAMP'));
        tb_temp.AddChampSupValeur('DH_LIBELLE',tb_tmp.GetString('DH_LIBELLE'));
        tb_tmp.free;
      end;
    end;

    // Gestion de la liste "Tri".
    s_orderby := Tb_param.Detail[idx].GetString('PMD_ORDERBY');
    while s_orderby <> '' do
    begin
      s_col := READTOKENPipe(s_orderby,',');
      s_ordre := copy(s_col,pos(' ',s_col)+1,length(s_col));
      s_col := copy(s_col,1,pos(' ',s_col)-1);

      tb_tmp := Tb_affiche.FindFirst(['DH_NOMCHAMP'],[s_col],FALSE);
      if tb_tmp <> nil then
      begin
        Obj_LstBxTri.Items.Add(tb_tmp.GetString('DH_LIBELLE'));
        tb_temp := tob.create('',Tb_tri,-1);
        tb_temp.AddChampSupValeur('DH_NOMCHAMP',tb_tmp.GetString('DH_NOMCHAMP'));
        tb_temp.AddChampSupValeur('DH_LIBELLE',tb_tmp.GetString('DH_LIBELLE'));
        tb_temp.AddChampSupValeur('DH_ORDRE',s_ordre);
      end;
    end;
  end;

  Tb_dispo.Detail.Sort('DH_LIBELLE');
  for idx:=0 to Tb_dispo.Detail.Count -1 do
  begin
    s_caption := Tb_dispo.Detail[idx].GetString('DH_LIBELLE');
    Obj_LstBxDispo.Items.Add(s_caption);
  end;

  if Obj_LstBxAff.Items.Count = 0 then
    Obj_LstBxAff.Items.Add('<<Toutes>>');

  Obj_LstBxDispo.ItemIndex := 0;
  Obj_LstBxAff.ItemIndex   := 0;
  Obj_LstBxTri.ItemIndex   := 0;
  Obj_LstBxTri.OnClick(Obj_LstBxTri);
  //Initialisation du Groupe Option
  Obj_EdtTxt.Text := Tb_varIN.Detail[2].GetString('PARAM_2');

  if Tb_varIN.Detail[0].GetString('PARAM_0') = 'UTILISAT' then
    Obj_GrpBxopt.Visible := FALSE;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.DoClickRdGrp(sender: Tobject);
var
 idx : integer;
 tb_tmp : TOB;
 s_col : string;
begin
  if (sender as THRadioGroup).name = 'RDGRP_VISUCHX' then
  begin
    case Obj_RdGrp.ItemIndex of
      0: AffichageEcran('COL');
      1: AffichageEcran('LIB');
    end;
  end
  else if (sender as THRadioGroup).name = 'RDGRP_ORDRE' then
  begin
    case Obj_RdGrpOrd.ItemIndex of
      0: begin;
           if Obj_LstBxTri.ItemIndex <> -1 then
           begin
             idx := Obj_LstBxTri.ItemIndex;
             s_col := tb_tri.Detail[idx].GetString('DH_NOMCHAMP');
             tb_tmp := tb_tri.FindFirst(['DH_NOMCHAMP'],[s_col],FALSE);
             if tb_tmp <> nil then
               tb_tmp.PutValue('DH_ORDRE','ASC');
           end;
         end;
      1: begin
           if Obj_LstBxTri.ItemIndex <> -1 then
           begin
             idx := Obj_LstBxTri.ItemIndex;
             s_col := tb_tri.Detail[idx].GetString('DH_NOMCHAMP');
             tb_tmp := tb_tri.FindFirst(['DH_NOMCHAMP'],[s_col],FALSE);
             if tb_tmp <> nil then
               tb_tmp.PutValue('DH_ORDRE','DESC');
           end;
         end;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.DoClickButton(sender: Tobject);
var
  idx,itm : integer;
  tb_tmp : tob;
  s_items : string;
  p_order : Pointer;
begin
  if (sender as TToolbarButton97).Name = 'BRIGHT' then
  begin
    // Déplace les champs de la liste "Colonnes disponibles" vers la liste "Afficher les colonnes suivantes".
    if Obj_LstBxDispo.ItemIndex <> -1 then
    begin
      if Obj_LstBxAff.Items.Strings[0] = '<<Toutes>>' then
        Obj_LstBxAff.Clear;

      idx := Obj_LstBxDispo.ItemIndex;
      itm := Obj_LstBxAff.items.Add(Obj_LstBxDispo.Items[idx]);

      Obj_LstBxAff.ItemIndex := itm;
      tb_tmp := tob.create('',Tb_affiche,-1);
      tb_tmp.AddChampSupValeur('DH_NOMCHAMP',tb_dispo.Detail[idx].GetString('DH_NOMCHAMP'));
      tb_tmp.AddChampSupValeur('DH_LIBELLE',tb_dispo.Detail[idx].GetString('DH_LIBELLE'));

      Obj_LstBxDispo.Items.Delete(idx);
      tb_dispo.Detail[idx].Free;

      if idx<Obj_LstBxDispo.Count-1 then
        Obj_LstBxDispo.ItemIndex := idx
      else
        Obj_LstBxDispo.ItemIndex := Obj_LstBxDispo.Count-1;
    end;
  end
  else if (sender as TToolbarButton97).Name = 'BLEFT' then
  begin
    // Déplace les champs de la liste "Afficher les colonnes suivantes" vers la liste "Colonnes disponibles".
    if (Obj_LstBxAff.ItemIndex <> -1) and (Obj_LstBxAff.Items.Strings[0] <> '<<Toutes>>') then
    begin
      idx := Obj_LstBxAff.ItemIndex;

      itm := Obj_LstBxDispo.items.Add(Obj_LstBxAff.Items[idx]);
      Obj_LstBxDispo.ItemIndex := itm;
      tb_tmp := tob.create('',Tb_dispo,-1);
      tb_tmp.AddChampSupValeur('DH_NOMCHAMP',Tb_affiche.Detail[idx].GetString('DH_NOMCHAMP'));
      tb_tmp.AddChampSupValeur('DH_LIBELLE',Tb_affiche.Detail[idx].GetString('DH_LIBELLE'));

      Obj_LstBxAff.Items.Delete(idx);
      if Obj_LstBxAff.Items.Count = 0 then
        Obj_LstBxAff.Items.Add('<<Toutes>>');

      Tb_affiche.Detail[idx].Free;

      if idx<Obj_LstBxAff.Count-1 then
        Obj_LstBxAff.ItemIndex := idx
      else
        Obj_LstBxAff.ItemIndex := Obj_LstBxAff.Count-1;
    end;
  end
  else if (sender as TToolbarButton97).Name = 'BUP' then
  begin
    // Déplace vers le haut le champs sélectionné dans la liste "Afficher les colonnes suivantes".
    idx := Obj_LstBxAff.ItemIndex;
    if idx > 0 then
    begin
      s_items  := Obj_LstBxAff.Items[idx];
      p_order := Tb_affiche.Detail[idx];

      Obj_LstBxAff.Items.Delete(idx);
      Tb_affiche.Detail.Delete(idx);

      Obj_LstBxAff.Items.Insert(idx-1, s_items);
      Tb_affiche.Detail.Insert(idx-1,p_order);

      Obj_LstBxAff.ItemIndex := idx-1;
    end;
  end
  else if (sender as TToolbarButton97).Name = 'BDOWN' then
  begin
    // Déplace vers le bas le champs sélectionné dans la liste "Afficher les colonnes suivantes".
    idx := Obj_LstBxAff.ItemIndex;
    if (idx+1) <= Obj_LstBxAff.Items.Count -1 then
    begin
      s_items  := Obj_LstBxAff.Items[idx];
      p_order := Tb_affiche.Detail[idx];

      Obj_LstBxAff.Items.Delete(idx);
      Tb_affiche.Detail.Delete(idx);

      Obj_LstBxAff.Items.Insert(idx+1, s_items);
      Tb_affiche.Detail.Insert(idx+1,p_order);

      Obj_LstBxAff.ItemIndex := idx+1;
    end;
  end
  else if (sender as TToolbarButton97).Name = 'BRIGHT_TRI' then
  begin
    // Déplace vers la droite le champ sélectionné dans la liste "Afficher les colonnes suivantes" vers la liste "Tri".
    if (Obj_LstBxAff.ItemIndex <> -1) and (Obj_LstBxAff.Items.Strings[0] <> '<<Toutes>>') then
    begin
      idx := Obj_LstBxAff.ItemIndex;
      tb_tmp := Tb_tri.FindFirst(['DH_NOMCHAMP'],[Tb_affiche.detail[idx].GetString('DH_NOMCHAMP')],FALSE);
      if tb_tmp = nil then
      begin
        itm := Obj_LstBxTri.items.Add(Obj_LstBxAff.Items[idx]);
        Obj_LstBxTri.ItemIndex := itm;

        tb_tmp := tob.create('',Tb_tri,-1);
        tb_tmp.AddChampSupValeur('DH_NOMCHAMP',Tb_affiche.Detail[idx].GetString('DH_NOMCHAMP'));
        tb_tmp.AddChampSupValeur('DH_LIBELLE',Tb_affiche.Detail[idx].GetString('DH_LIBELLE'));
        case Obj_RdGrpOrd.ItemIndex of
          0: tb_tmp.AddChampSupValeur('DH_ORDRE','ASC');
          1: tb_tmp.AddChampSupValeur('DH_ORDRE','DESC');
        end;
      end;
    end;
  end
  else if (sender as TToolbarButton97).Name = 'BLEFT_TRI' then
  begin
    if Obj_LstBxTri.ItemIndex <> -1 then
    begin
      idx := Obj_LstBxTri.ItemIndex;
      Obj_LstBxTri.Items.Delete(idx);
      Tb_tri.Detail[idx].Free;
      Obj_LstBxTri.ItemIndex := 0;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.AffichageEcran(s_mode: String);
var
  idx : integer;
  s_caption : string;
begin
  if Obj_LstBxAff.Items.Strings[0] <> '<<Toutes>>' then
    Obj_LstBxAff.Clear;

  Obj_LstBxDispo.Clear;
  Obj_LstBxTri.Clear;

  if s_mode = 'COL' then
  begin
    Tb_dispo.Detail.Sort('DH_NOMCHAMP');
    for idx:=0 to Tb_dispo.Detail.Count -1 do
    begin
      s_caption := Tb_dispo.Detail[idx].GetString('DH_NOMCHAMP');
      Obj_LstBxDispo.Items.Add(s_caption);
    end;

    for idx:=0 to Tb_affiche.Detail.Count -1 do
    begin
      s_caption := Tb_affiche.Detail[idx].GetString('DH_NOMCHAMP');
      Obj_LstBxAff.Items.Add(s_caption);
    end;

    for idx:=0 to Tb_tri.Detail.Count -1 do
    begin
      s_caption := Tb_tri.Detail[idx].GetString('DH_NOMCHAMP');
      Obj_LstBxTri.Items.Add(s_caption);
    end;
  end
  else
  begin
    Tb_dispo.Detail.Sort('DH_LIBELLE');
    for idx:=0 to Tb_dispo.Detail.Count -1 do
    begin
      s_caption := Tb_dispo.Detail[idx].GetString('DH_LIBELLE');
      Obj_LstBxDispo.Items.Add(s_caption);
    end;

    for idx:=0 to Tb_affiche.Detail.Count -1 do
    begin
      s_caption := Tb_affiche.Detail[idx].GetString('DH_LIBELLE');
      Obj_LstBxAff.Items.Add(s_caption);
    end;

    for idx:=0 to Tb_tri.Detail.Count -1 do
    begin
      s_caption := Tb_tri.Detail[idx].GetString('DH_LIBELLE');
      Obj_LstBxTri.Items.Add(s_caption);
    end;
  end;

  Obj_LstBxDispo.ItemIndex := 0;
  Obj_LstBxAff.ItemIndex := 0;
  Obj_LstBxTri.ItemIndex := 0;
  Obj_LstBxTri.OnClick(Obj_LstBxTri);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.OnUpdate ;
var
  idx : integer;
  s_select,s_orderby,s_sql : string;
  tb_tmp : tob;
begin
  Inherited ;
  s_select  := '';
  s_orderby := '';

  for idx := 0 to Tb_affiche.Detail.Count -1 do
  begin
    s_select := s_select+Tb_affiche.Detail[idx].GetString('DH_NOMCHAMP')+',';
  end;
  // retrait de la virgule en fin de chaine
  if s_select <> '' then
    s_select := copy(s_select,1,length(s_select)-1)
  else
    s_select := '*';

  for idx := 0 to Tb_tri.Detail.Count -1 do
  begin
    s_orderby := s_orderby+Tb_tri.Detail[idx].GetString('DH_NOMCHAMP')+
                 ' '+Tb_tri.Detail[idx].GetString('DH_ORDRE')+',';
  end;
  if s_orderby <> '' then
    s_orderby := copy(s_orderby,1,length(s_orderby)-1);

  //maj de la table paramdonnees
  Tb_param.ClearDetail;
  tb_tmp := tob.create('PARAMDONNEES',Tb_param,-1);
  tb_tmp.PutValue('PMD_GRPNOM',Tb_varIN.Detail[1].GetString('PARAM_1'));
  tb_tmp.PutValue('PMD_NOMTABLE',Tb_varIN.Detail[0].GetString('PARAM_0'));
  tb_tmp.PutValue('PMD_SELECT',s_select);
  tb_tmp.PutValue('PMD_ORDERBY',s_orderby);
  tb_tmp.InsertOrUpdateDB;

  //mise du libelle pour tous selon le groupe de données.
  s_sql := 'UPDATE PARAMDONNEES SET PMD_LIBELLE="'+Obj_EdtTxt.Text+'" '+
           'WHERE PMD_GRPNOM = "'+Tb_varIN.Detail[1].GetString('PARAM_1')+'"';
  ExecuteSQL(s_sql);

  TFVierge(ecran).retour := 'mbOK';
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.DoClickOnLstBxTri(sender: Tobject);
var
 idx : integer;
 s_col : string;
begin
  if Obj_LstBxTri.ItemIndex <> -1 then
  begin
    idx := Obj_LstBxTri.ItemIndex;
    s_col := tb_tri.Detail[idx].GetString('DH_ORDRE');
    if s_col = 'ASC' then
      Obj_RdGrpOrd.ItemIndex := 0
    else
      Obj_RdGrpOrd.ItemIndex := 1;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PARAM_GRPTRAVAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PARAM_GRPTRAVAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PARAM_GRPTRAVAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PARAM_GRPTRAVAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PARAM_GRPTRAVAIL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PARAM_GRPTRAVAIL ] ) ; 
end.
