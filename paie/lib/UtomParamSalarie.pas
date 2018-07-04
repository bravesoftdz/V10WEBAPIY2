{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARAMSALARIE (PARAMSALARIE)
Mots clefs ... : TOM;PARAMSALARIE
*****************************************************************
PT1   02/07/2007 FC V_72 : FQ 14431 Doublon dans la liste des valeurs
PT2   04/09/2007 GGU v_8 : FQ 14677 Probleme de filtre sur la tablette des tables dynamiques
PT3   05/11/2007 FC V_80 : FQ 14908 Journal événement
PT4   02/01/2008 FC V_81 : FQ 14436 Contrôler que le code table existe
PT6   16/04/2008 GGU V_81 : FQ 15362 problème d'affichage des tables dynamiques dans le paramétrage d'un élément dynamique
PT7  : 17/04/2008 GGU V81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
PT13  08/09/2008 FC Rajout d'un controle sur le champ zone salarié dans le cas saisie historique salarié
}
Unit UtomParamSalarie ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
     HDB,Fe_Main,
{$else}
     eFiche,
     eFichList,
     MaineAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     P5Util,
     PGOutils,
     PGOutils2,
     UTob,
     LookUp,
     HTB97,
     PAIETOM,P5Def ; //PT3
Type
  TOM_PARAMSALARIE = Class (PGTOM) //PT3 class TOM devient PGTOM
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
      TypeSaisie : String;
      ListeEtab, ListeConv: string; //PT1
      HadSearchEtabConv: Boolean; //PT1
      Trace: TStringList;                //PT3
      DerniereCreate: string;            //PT3
      LeStatut:TDataSetState;            //PT3
      procedure TableElipsisClick(Sender: TObject); //PT1
      procedure RechercheEtabConv;//PT1
      procedure AccesTableDyna(Sender: TObject);
    end ;

Implementation

procedure TOM_PARAMSALARIE.OnNewRecord ;
begin
  Inherited ;
  SetField('PPP_PGTYPEINFOLS',TypeSaisie);
  If TypeSaisie = 'ZLS' then SetField('PPP_LIENASSOC','PGCOMBOZONELIBRE');
end ;

procedure TOM_PARAMSALARIE.OnDeleteRecord ;
begin
  //DEB PT3
  Trace := TStringList.Create ;
  Trace.Add('SUPPRESSION ELEMENT DYNAMIQUE '+GetField('PPP_PGINFOSMODIF')+' '+ GetField('PPP_LIBELLE'));
  CreeJnalEvt('003','093','OK',nil,nil,Trace);
  FreeAndNil (Trace);
  //FIN PT3
end ;

procedure TOM_PARAMSALARIE.OnUpdateRecord ;
var NoDossier,Suffix,Pred : String;
    Q : TQuery;
    CEG, STD, DOS : Boolean;
begin
  Inherited ;
  //DEB PT13
  if TypeSaisie = 'SAL' then
  begin
    if RechDom('PGCHAMPSAL',GetField('PPP_PGINFOSMODIF'),false) = '' then
    begin
      PGIBox('Le code Zone salarié saisi n''existe pas dans le paramétrage');
      LastError := 1;
      SetFocusControl('PPP_PGINFOSMODIF');
      Exit;
    end;
  end;
  //FIN PT13
  AccesPredefini('TOUS', CEG, STD, DOS);
  Pred := GetField('PPP_PREDEFINI');
  if (Pred = 'CEG') and (CEG = FALSE) then
  begin
    PGIBox('Vous ne pouvez pas créer d''élément prédéfini CEGID', 'Accès refusé');
    LastError := 1;
    Pred := '';
    SetControlProperty('PEL_PREDEFINI', 'Value', Pred);
    SetFocusControl('PPP_PREDEFINI');
    Exit;
  end;
  if (Pred = 'STD') and (STD = FALSE) then
  begin
    PGIBox('Vous ne pouvez pas créer d''élément prédéfini Standard', 'Accès refusé');
    LastError := 1;
    Pred := '';
    SetControlProperty('PEL_PREDEFINI', 'Value', Pred);
    SetFocusControl('PPP_PREDEFINI');
    Exit;
  end;
  if Pred = 'DOS' then NoDossier := PgRendNoDossier()
  else NoDossier := '000000';
  Setfield('PPP_NODOSSIER',NoDossier);
  If TypeSaisie = 'SAL' then
  begin
      Suffix := GetField('PPP_PGINFOSMODIF');
      Suffix := Copy(Suffix,5,Length(Suffix));
      Q := OpenSQL('SELECT * FROM PAIEPARIM WHERE PAI_SUFFIX="'+Suffix+'"',True);
      If Not Q.Eof then
      begin
        SetField('PPP_LIENASSOC',Q.FindField('PAI_LIENASSOC').AsString);
        SetField('PPP_PGTYPEDONNE',Q.FindField('PAI_LETYPE').AsString);
      end;
      Ferme(Q);
  end
  else
  begin
    If getfield('PPP_PGTYPEDONNE') = 'T' then
    begin
      SetField('PPP_LIENASSOC','PGCOMBOZONELIBRE');
      If ExisteSQL('SELECT PPP_PGINFOSMODIF FROM PARAMSALARIE WHERE PPP_LIENASSOC="PGCOMBOZONELIBRE" AND '+
      'PPP_CODTABL="'+GetField('PPP_CODTABL')+'" AND PPP_PGINFOSMODIF<>"'+GetField('PPP_PGINFOSMODIF')+'"') then
      begin
        LastError := 1;
        PGIBox('Validation impossible car il existe déja une zone libre avec cette table dynamique',Ecran.Caption);
      end;
      //DEB PT4
      if not ExisteSQL('SELECT PTE_CODTABL FROM TABLEDIMENT WHERE PTE_CODTABL="' + GetField('PPP_CODTABL') + '"') then
      begin
        LastError := 1;
        PGIBox('Validation impossible car le code table dynamique saisi n''existe pas',Ecran.Caption);
      end;
      //FIN PT4
    end
    else SetField('PPP_LIENASSOC','');
  end;

  //DEB PT3
  if (DS.State = dsinsert) then
    DerniereCreate := GetField('PPP_PGINFOSMODIF');
  LeStatut := DS.State;
  //FIN PT3
end ;

procedure TOM_PARAMSALARIE.OnAfterUpdateRecord ;
var
  even: boolean;
  LaTable,lecode, LeLibelle : String;
begin
  //DEB PT3
  LaTable := 'PARAMSALARIE';
  LeCode := 'PPP_PGINFOSMODIF';
  LeLibelle := 'PPP_LIBELLE';
  Trace := TStringList.Create ;
  even := IsDifferent(dernierecreate,Latable,LeCode,LeLibelle,Trace,TFFicheListe(Ecran),LeStatut);
  FreeAndNil (Trace);
  //FIN PT3
end ;

procedure TOM_PARAMSALARIE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMSALARIE.OnLoadRecord ;
begin
  Inherited ;
  If TypeSaisie = 'SAL' then TFFiche(Ecran).Caption := 'Saisie historique'
  else TFFiche(Ecran).Caption := 'Saisie zone libre';
  UpdateCaption(TFFiche(Ecran));
  if (DS.State in [dsInsert]) then
  else
    DerniereCreate := ''; //PT3
end ;

procedure TOM_PARAMSALARIE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  If F.FieldName='PPP_PGTYPEDONNE' then
  begin
    If typeSaisie = 'ZLS' then
    begin
      SetControlEnabled('PPP_CODTABL',GetField('PPP_PGTYPEDONNE')='T');
      SetControlEnabled('BNEWTABLEDYNA',GetField('PPP_PGTYPEDONNE')='T');
    end;
  end;
  If F.FieldName='PPP_TYPENIVEAU' then
  begin
    If getField('PPP_TYPENIVEAU') <> 'SAL' then
    begin
      SetControlEnabled('PPP_PGTHEMESALARIE',True);
      If GetField('PPP_PGTHEMESALARIE') <> '' then SetField('PPP_PGTHEMESALARIE','');
    end
    else SetControlEnabled('PPP_PGTHEMESALARIE',True);
  end;
end ;

procedure TOM_PARAMSALARIE.OnArgument ( S: String ) ;
var
//DEB PT1
{$IFNDEF EAGLCLIENT}
  Defaut: THDBEdit;
{$ELSE}
  Defaut: THEdit;
{$ENDIF}
//FIN PT1
  Btn: TToolBarButton97;
begin
  Inherited ;
  TypeSaisie := ReadTokenPipe(S,';');

  //DEB PT1
{$IFNDEF EAGLCLIENT}
  Defaut:=ThDBEdit(getcontrol('PPP_CODTABL'));
{$ELSE}
  Defaut:=ThEdit(getcontrol('PPP_CODTABL'));
{$ENDIF}
  If Defaut<>nil then Defaut.OnElipsisClick := TableElipsisClick;
  HadSearchEtabConv := False;
  //FIN PT1

  Btn := TToolBarButton97(GetControl('BNEWTABLEDYNA'));
  if Btn <> nil then Btn.OnClick := AccesTableDyna;
end ;

//DEB PT1
procedure TOM_PARAMSALARIE.TableElipsisClick(Sender: TObject);
var
  StWhere : String;
  Q : TQuery;
begin
  RechercheEtabConv;
  StWhere := ' ##PTE_PREDEFINI## '
    + ' AND PTE_DTVALID = (SELECT MAX(Z.PTE_DTVALID) '
    + '                        FROM TABLEDIMENT Z '
//PT7    + '                       WHERE Z.PTE_CODTABL=TD.PTE_CODTABL '
                          + ' WHERE ##Z.PTE_PREDEFINI## ' //PT7
                          + ' AND TD.PTE_CODTABL = Z.PTE_CODTABL '
                          + ' AND TD.PTE_PREDEFINI = Z.PTE_PREDEFINI '
                          + ' AND TD.PTE_NODOSSIER = Z.PTE_NODOSSIER '
                          + ' AND TD.PTE_NIVSAIS = Z.PTE_NIVSAIS '
                          + ' AND TD.PTE_DTVALID<="'+UsDateTime(Date())+'")'
//PT7    + '                         AND TD.PTE_DTVALID<="' + UsDateTime(Date()) + '")'
    + ' AND PTE_NATURETABLE = "COD" ';
  { PT2 On filtre aussi en fonction du prédefini de l'élément dynamique }
  if GetField('PPP_PREDEFINI') = 'CEG' then
  begin
    StWhere := StWhere + ' AND (PTE_PREDEFINI = "CEG")';
  end else if GetField('PPP_PREDEFINI') = 'STD' then
  begin
    StWhere := StWhere + ' AND (PTE_PREDEFINI <> "DOS")';
  end else if GetField('PPP_PREDEFINI') = 'DOS' then
//PT6 : On ne fait le filtre sur les conventions et établissement que pour les saisies de niveau Dossier
  begin
    StWhere := StWhere + ' AND (   PTE_NIVSAIS = "GEN" '
      + '      or PTE_PREDEFINI = "DOS" '
      + '      or (  (PTE_NIVSAIS = "CON" and PTE_VALNIV in (' + ListeConv + '"000")) '
      + '         or (PTE_NIVSAIS = "ETB" and PTE_VALNIV in (' + ListeEtab + '"...")) '
      + '         ) '
      + '      ) ';
  end;

  LookupList(Sender as TControl
    , 'Tables dynamiques'
    , 'TABLEDIMENT TD'
    , 'DISTINCT PTE_CODTABL'
    , 'PTE_LIBELLE'
    , StWhere
    , 'TD.PTE_CODTABL', True, -1);

  if GetControlText('PPP_CODTABL') <> '' then
  begin
    Q := Opensql('SELECT PTE_LIBELLE FROM TABLEDIMENT WHERE PTE_CODTABL = "' + GetControlText('PPP_CODTABL') + '"', true);
    if not Q.eof then
      SetControlText('LIBTABLE', Q.findfield('PTE_LIBELLE').AsString);
    Ferme(Q);
  end;
end;

procedure TOM_PARAMSALARIE.RechercheEtabConv;
var
  Qry: TQuery;
  Requete: string;
begin
if not HadSearchEtabConv then
   begin
   Requete:= 'SELECT ET_ETABLISSEMENT, ET_NODOSSIER, ETB_CONVENTION,'+
             ' ETB_CONVENTION1, ETB_CONVENTION2'+
             ' FROM ETABLISS'+
             ' LEFT JOIN ETABCOMPL ON'+
             ' ET_ETABLISSEMENT = ETB_ETABLISSEMENT WHERE'+
             ' (ET_NODOSSIER="000000" OR'+
             ' ET_NODOSSIER="'+V_PGI.NoDossier+'")';
   Qry:= opensql (Requete, true);
   while not Qry.EOF do
         begin
         if pos (Qry.findfield ('ET_ETABLISSEMENT').asstring, ListeEtab)=0 then
            ListeEtab:= ListeEtab+'"'+Qry.findfield ('ET_ETABLISSEMENT').asstring+'", ';
         if pos (Qry.findfield ('ETB_CONVENTION').asstring, ListeConv)=0 then
            ListeConv:= ListeConv+'"'+Qry.findfield ('ETB_CONVENTION').asstring+'", ';
         if pos (Qry.findfield ('ETB_CONVENTION1').asstring, ListeConv)=0 then
            ListeConv:= ListeConv+'"'+Qry.findfield ('ETB_CONVENTION1').asstring+'", ';
         if pos (Qry.findfield ('ETB_CONVENTION2').asstring, ListeConv)=0 then
            ListeConv:= ListeConv+'"'+Qry.findfield ('ETB_CONVENTION2').asstring+'", ';
         Qry.next;
         end;
   ferme(Qry);
   HadSearchEtabConv := True;
   end;
end;
//FIN PT1

procedure TOM_PARAMSALARIE.AccesTableDyna(Sender: TObject);
var
  St : String;
  Q : TQuery;
begin
  St := 'SELECT PTE_CODTABL,PTE_PREDEFINI,PTE_NODOSSIER,PTE_DTVALID,PTE_NIVSAIS,PTE_VALNIV FROM TABLEDIMENT ' +
    ' WHERE ##PTE_PREDEFINI## PTE_CODTABL = "' + GetField('PPP_CODTABL') + '" ORDER BY PTE_DTVALID DESC';
  Q := OpenSQL(St, true);
  if Q.Eof then
    AglLanceFiche('PAY', 'PGTABLESDYNA', '', '', 'ACTION=CREATION;;;;')
  else
    AglLanceFiche('PAY', 'PGTABLESDYNA', '', Q.FindField('PTE_CODTABL').AsString+';'+Q.FindField('PTE_PREDEFINI').AsString+';'+Q.FindField('PTE_NODOSSIER').AsString+';'+Q.FindField('PTE_DTVALID').AsString+';'+Q.FindField('PTE_NIVSAIS').AsString+';'+Q.FindField('PTE_VALNIV').AsString,'ACTION=MODIFICATION;;;;');
  Ferme(Q);
end;

Initialization
  registerclasses ( [ TOM_PARAMSALARIE ] ) ;
end.

