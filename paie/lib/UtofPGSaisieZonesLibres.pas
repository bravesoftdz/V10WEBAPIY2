{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSAISIEZL ()
Mots clefs ... : TOF;PGSAISIEZL
*****************************************************************
PT1   02/01/2008 FC V810 FQ 15085 Si libellé, la valeur saisie ne pointe pas sur l'élément auquel elle appartient
PT2   18/04/2008 GGU V_81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
}
Unit UtofPGSaisieZonesLibres ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
 //    uTob,
     MainEAGL,
{$ENDIF}
     forms,
     HSysMenu,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Grids,
     graphics,
     Utob,
     Vierge,
     HTB97,
     UTOF ;

Type
  TOF_PGSAISIEZL = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    TypeSaisie,LeSalarie,Etabliss,LeTheme,CodePop,TypePop : String;
    procedure ChangeComboEltComp(Sender : Tobject);
    procedure ClickBtEltCompl(Sender : Tobject);
    procedure AfficheEltCompl;
    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure GrilleEltComplClick(Sender : TObject);
  end ;

Implementation

Uses
  PGTablesDyna //PT2
  ;

procedure TOF_PGSAISIEZL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGSAISIEZL.OnArgument (S : String ) ;
var Q : TQuery;
    Grille : THGrid;
    Btn : TToolBarButton97;
    Combo : THValComboBox;
begin
  Inherited ;
   SetControlText('CZONELIBRE','OUI');
   Etabliss := '';CodePop :='';TypePop:='';
   TypeSaisie := ReadTokenPipe(S,';');
   LeSalarie := ReadTokenPipe(S,';');
   LeTheme := ReadTokenPipe(S,';');
   If TypeSaisie = 'SAL' then
   begin
     Q := OpenSQL('SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="'+LeSalarie+'"',True);
     Etabliss := Q.FindField('PSA_ETABLISSEMENT').AsString;
    Ferme(Q);
    TFVierge(Ecran).Caption := 'Saisie éléments dynamiques pour le salarié';
   end
   else
   if TypeSaisie = 'ETB' then
   begin
    Etabliss := LeSalarie;
    LeTheme := '';
    LeSalarie := '';
    TFVierge(Ecran).Caption := 'Saisie éléments dynamiques pour l''établissement '+RechDom('TTETABLISSEMENT',Etabliss,False);
   end
   else
   if TypeSaisie = 'POP' then
   begin
    CodePop := LeSalarie;
    TypePop := LeTheme;
    LeTheme := '';
    LeSalarie := '';
    TFVierge(Ecran).Caption := 'Saisie éléments dynamiques pour la population';
   end
   else
   if TypeSaisie = 'DOS' then
   begin
    LeTheme := '';
    LeSalarie := '';
    TFVierge(Ecran).Caption := 'Saisie éléments dynamiques pour le dossier : '+V_PGI.NoDossier;
   end;
   UpdateCaption(TFVierge(Ecran));
   SetControlText('CTHEMEELTCOMPL',LeTheme);
   AfficheEltCompl;
  Grille := THGrid(GetControl('GELTCOMPL'));
  If Grille <> Nil then
  begin
    Grille.OnDblClick := GrilleEltComplClick;
    Grille.GetCellCanvas := GrilleGetCellCanvas;
  end;
  Btn := TToolBarbutton97(GetControl('BRECHELTCOMP'));
  If Btn <> Nil then Btn.OnClick := ClickBtEltCompl;
  combo := THValComboBox(GetControl('CTHEMEELTCOMPL'));
  If combo <> Nil then combo.OnClick := ChangeComboEltComp;
  combo := THValComboBox(GetControl('CZONELIBRE'));
  If combo <> Nil then combo.OnClick := ChangeComboEltComp;
end ;

procedure TOF_PGSAISIEZL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGSAISIEZL.ChangeComboEltComp(Sender : Tobject);
begin
  AfficheEltCompl;
end;

procedure TOF_PGSAISIEZL.ClickBtEltCompl(Sender : Tobject);
begin
  AfficheEltCompl;
end;

procedure TOF_PGSAISIEZL.AfficheEltCompl;
var Q : TQuery;
    TobParamSal,TobGrille,TG : Tob;
    i : Integer;
    Libelle,LeChamp,Valeur,Letype : String;
    GrilleEC : THGrid;
    HMTrad : THSystemMenu;
    WhereLib,WhereTheme,WhereInfo : String;
    Theme,ThemePrec,Niveau,Img,Tablette : String;
    DateApplic : TDateTime;
begin
  DateApplic := iDate1900;
  If TypeSaisie = 'ETB' then WhereInfo := 'PHD_ETABLISSEMENT="'+Etabliss+'" AND '
  else If TypeSaisie = 'POP' then WhereInfo := 'PHD_CODEPOP="'+CodePop+'" AND PHD_POPULATION="'+TypePop+'" AND '
  else If TypeSaisie = 'DOS' then WhereInfo := 'PHD_ETABLISSEMENT="" AND PHD_SALARIE="" AND PHD_CODEPOP="" AND '
  else WhereInfo := '';
  If GetControlText('CTHEMEELTCOMPL') <> '' then WhereTheme := ' AND PPP_PGTHEMESALARIE="'+GetControlText('CTHEMEELTCOMPL')+'"'
  else Wheretheme := '';
  If GetControlText('ELIBELTCOMPL') <> '' then WhereLib := ' AND PPP_LIBELLE LIKE "%'+GetControlText('ELIBELTCOMPL')+'%"'
  else WhereLib := '';
  GrilleEC := THGrid(GetControl('GELTCOMPL'));
  For i := 1 to GrilleEC.RowCount do
  begin
    GrilleEC.Rows[i].Clear;
  end;
  GrilleEC.RowCount := 2;
  GrilleEC.FixedRows := 1;
  If GetControlText('CZONELIBRE') = 'OUI' then
  begin
    Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_TYPENIVEAU="'+TypeSaisie+'" AND PPP_PGTYPEINFOLS="ZLS"'+WhereTheme+WhereLib+
    ' AND PPP_PGINFOSMODIF IN (SELECT PHD_PGINFOSMODIF FROM PGHISTODETAIL WHERE '+WhereInfo+
    'PHD_DATEAPPLIC<="'+UsDateTime(V_PGI.DateEntree)+'") ORDER BY PPP_PGTHEMESALARIE',True);
    TobParamSal := Tob.Create('parametrage salarie',Nil,-1);
    TobParamSal.LoadDetailDB('Parametrage', '', '', Q, FALSE, False);
    ferme(Q);
    Theme := '';
    ThemePrec := '';
    TobGrille := Tob.Create('remplir grille',Nil,-1);
      For i := 0 to TobParamSal.Detail.Count -1 do
      begin
          Theme := TobParamSal.Detail[i].GetValue('PPP_PGTHEMESALARIE');
          Niveau := TobParamSal.Detail[i].GetValue('PPP_TYPENIVEAU');
          Tablette := TobParamSal.Detail[i].GetValue('PPP_CODTABL');
          If (Theme='') and (i = 0) then
          begin
            TG := Tob.Create('ligne grille',TobGrille,-1);
            TG.AddChampSupValeur('NIVEAU','Aucun thème');
            TG.AddChampSupValeur('ZONE','Aucun thème');
            TG.AddChampSupValeur('VALEURAFFICHE','');
            TG.AddChampSupValeur('DD','');
            TG.AddChampSupValeur('LETYPE','THEME');
            TG.AddChampSupValeur('CHAMP','');
          end
          else If (Theme <> '') and (Theme <> ThemePrec) then
          begin
            TG := Tob.Create('ligne grille',TobGrille,-1);
            TG.AddChampSupValeur('NIVEAU',RechDom('PGTHEMESAL',Theme,False));
            TG.AddChampSupValeur('ZONE',RechDom('PGTHEMESAL',Theme,False));
            TG.AddChampSupValeur('VALEURAFFICHE','');
            TG.AddChampSupValeur('DD','');
            TG.AddChampSupValeur('ETAT','');
            TG.AddChampSupValeur('LETYPE','THEME');
            TG.AddChampSupValeur('CHAMP','');
          end;
          LeChamp := TobParamSal.Detail[i].GetValue('PPP_PGINFOSMODIF');
          If LeChamp = 'PSA_SITUATIONFAMI' then LeChamp := 'PSA_SITUATIONFAMIL';
          Libelle := TobParamSal.Detail[i].GetValue('PPP_LIBELLE');
          Letype := TobParamSal.Detail[i].GetValue('PPP_PGTYPEDONNE');
          Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE '+WhereInfo+
          'AND PHD_PGINFOSMODIF="'+LeChamp+'" AND PHD_DATEAPPLIC<="'+UsDateTime(V_PGI.DateEntree)+'" ORDER BY PHD_DATEAPPLIC DESC',True);
          If Not Q.Eof then
          begin
            Q.First;
            Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
            DateApplic := Q.FindField('PHD_DATEAPPLIC').AsDateTime;
            If Niveau <> 'SAL' then Img := '#ICO#91'
            else Img := '#ICO#44';
          end
          else
          begin
            Valeur := 'Aucune valeur';
            Img := '#ICO#43';
          end;
          Ferme(Q);
          TG := Tob.Create('ligne grille',TobGrille,-1);
          TG.AddChampSupValeur('NIVEAU',RechDom('PGNIVEAUAVDOS',Niveau,False));
          TG.AddChampSupValeur('ZONE',Libelle);
          If LeType = 'T' then
          begin
            //PT2 si LeSalarie = '', on ramène uniquement les lignes Toutes conventions, tous etabs ou générales des tables dynamiques
            Valeur := GetLibelleZLTableDyna(DateApplic, Tablette, LeSalarie, Valeur);
{//PT2
            Q := OpenSQL('SELECT ##TOP 1## PTD_LIBELLECODE FROM TABLEDIMDET WHERE PTD_DTVALID<="'+UsDateTime(DateApplic)+'"'+
            ' AND PTD_CODTABL="'+Tablette+'" AND PTD_VALCRIT1="'+Valeur+'" ORDER BY PTD_DTVALID DESC',True);
            If Not Q.Eof then Valeur := Q.FindField('PTD_LIBELLECODE').AsString;
            Ferme(Q);
}
          end;
          TG.AddChampSupValeur('VALEURAFFICHE',Valeur);
          TG.AddChampSupValeur('DD',DateApplic);
          TG.AddChampSupValeur('ETAT',Img);
          TG.AddChampSupValeur('LETYPE',Letype);
          TG.AddChampSupValeur('CHAMP',LeChamp);
          ThemePrec := Theme;
      end;
  end
  else
  begin
    Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE PPP_TYPENIVEAU="'+TypeSaisie+'" AND PPP_PGTYPEINFOLS="ZLS"'+WhereTheme+WhereLib+' '+
    'AND PPP_PGINFOSMODIF NOT IN (SELECT PHD_PGINFOSMODIF FROM PGHISTODETAIL WHERE '+WhereInfo+
    'AND PHD_DATEAPPLIC<="'+UsDateTime(V_PGI.DateEntree)+'") ORDER BY PPP_PGTHEMESALARIE',True);
    TobParamSal := Tob.Create('parametrage salarie',Nil,-1);
    TobParamSal.LoadDetailDB('Parametrage', '', '', Q, FALSE, False);
    ferme(Q);
    Theme := '';
    ThemePrec := '';
    TobGrille := Tob.Create('remplir grille',Nil,-1);
      For i := 0 to TobParamSal.Detail.Count -1 do
      begin
          Theme := TobParamSal.Detail[i].GetValue('PPP_PGTHEMESALARIE');
          Niveau := TobParamSal.Detail[i].GetValue('PPP_TYPENIVEAU');
          Tablette := TobParamSal.Detail[i].GetValue('PPP_CODTABL');
          If (Theme='') and (i = 0) then
          begin
            TG := Tob.Create('ligne grille',TobGrille,-1);
            TG.AddChampSupValeur('NIVEAU','Aucun thème');
            TG.AddChampSupValeur('ZONE','Aucun thème');
            TG.AddChampSupValeur('VALEURAFFICHE','');
            TG.AddChampSupValeur('DD','');
            TG.AddChampSupValeur('LETYPE','THEME');
            TG.AddChampSupValeur('CHAMP','');
          end
          else If (Theme <> '') and (Theme <> ThemePrec) then
          begin
            TG := Tob.Create('ligne grille',TobGrille,-1);
            TG.AddChampSupValeur('NIVEAU','');
            TG.AddChampSupValeur('ZONE',RechDom('PGTHEMESAL',Theme,False));
            TG.AddChampSupValeur('VALEURAFFICHE','');
            TG.AddChampSupValeur('DD','');
            TG.AddChampSupValeur('ETAT','');
            TG.AddChampSupValeur('LETYPE','THEME');
            TG.AddChampSupValeur('CHAMP','');
          end;
          LeChamp := TobParamSal.Detail[i].GetValue('PPP_PGINFOSMODIF');
          If LeChamp = 'PSA_SITUATIONFAMI' then LeChamp := 'PSA_SITUATIONFAMIL';
          Libelle := TobParamSal.Detail[i].GetValue('PPP_LIBELLE');
          Letype := TobParamSal.Detail[i].GetValue('PPP_PGTYPEDONNE');
          Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE '+WhereInfo+
          'AND PHD_PGINFOSMODIF="'+LeChamp+'" AND PHD_DATEAPPLIC<="'+UsDateTime(V_PGI.DateEntree)+'" ORDER BY PHD_DATEAPPLIC DESC',True);
          If Not Q.Eof then
          begin
            Q.First;
            Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
            DateApplic := Q.FindField('PHD_DATEAPPLIC').AsDateTime;
            If Niveau <> 'SAL' then Img := '#ICO#91'
            else Img := '#ICO#44';
          end
          else
          begin
            Valeur := 'Aucune valeur';
            Img := '#ICO#43';
          end;
          Ferme(Q);
          TG := Tob.Create('ligne grille',TobGrille,-1);
          TG.AddChampSupValeur('NIVEAU',RechDom('PGNIVEAUAVDOS',Niveau,False));
          TG.AddChampSupValeur('ZONE',Libelle);
          If LeType = 'T' then
          begin
            //PT2 si LeSalarie = '', on ramène uniquement les lignes Toutes conventions, tous etabs ou générales des tables dynamiques
            Valeur := GetLibelleZLTableDyna(DateApplic, Tablette, LeSalarie, Valeur);
{//PT2
            Q := OpenSQL('SELECT ##TOP 1## PTD_LIBELLECODE FROM TABLEDIMDET WHERE PTD_DTVALID<="'+UsDateTime(DateApplic)+'"'+
            ' AND PTD_CODTABL="'+Tablette+'" AND PTD_VALCRIT1="'+Valeur+'" ORDER BY PTD_DTVALID DESC',True);
            If Not Q.Eof then Valeur := Q.FindField('PTD_LIBELLECODE').AsString;
            Ferme(Q);
}
          end;
          TG.AddChampSupValeur('VALEURAFFICHE',Valeur);
          TG.AddChampSupValeur('DD',DateApplic);
          TG.AddChampSupValeur('ETAT',Img);
          TG.AddChampSupValeur('LETYPE',Letype);
          TG.AddChampSupValeur('CHAMP',LeChamp);
          ThemePrec := Theme;
      end;

  end;
  If TobGrille.Detail.Count > 0 then  GrilleEC.RowCount := TobGrille.Detail.Count + 1
  else GrilleEC.RowCount := 2;
  GrilleEC.ColCount := 7;
  GrilleEC.ColWidths[0] := -1;
  GrilleEC.ColWidths[1] := 200;
  GrilleEC.ColWidths[2] := 200;
  GrilleEC.ColWidths[3] := 150;
  GrilleEC.ColWidths[4] := -1;
  GrilleEC.ColWidths[5] := -1;
  GrilleEC.ColWidths[6] := -1;
  GrilleEC.FixedCols := 1;
  GrilleEC.ColFormats[2] := '';
  GrilleEC.ColAligns[2] := taCenter;
  GrilleEC.ColAligns[3] := taCenter;
  GrilleEC.ColFormats[3] := ShortDateFormat;
  If TobGrille.Detail.Count > 0 then TobGrille.PutGridDetail(GrilleEC,False,False,'NIVEAU;ZONE;VALEURAFFICHE;DD;ETAT;LETYPE;CHAMP',False)
  else
  begin
    GrilleEC.cellValues[1,1] := '<<Aucun élément>>';
  end;
  HMTrad.ResizeGridColumns(GrilleEC) ;
  TobGrille.Free;
  TobParamSal.free;
end;

procedure TOF_PGSAISIEZL.GrilleEltComplClick(Sender : TObject);
var GrilleEC : THGrid;
    Libelle,Tablette,Salarie,Ret,LeType,Champ : String;
    Ligne : Integer;
    Img,Valeur,Niveau : String;
    DateApplic : TDateTime;
    Q  : TQuery;
    WhereInfo,Code : String;
begin
  If TypeSaisie = 'ETB' then WhereInfo := 'PHD_ETABLISSEMENT="'+Etabliss+'" AND '
  else If TypeSaisie = 'POP' then WhereInfo := 'PHD_CODEPOP="'+CodePop+'" AND PHD_POPULATION="'+TypePop+'" AND '
  else If TypeSaisie = 'DOS' then WhereInfo := 'PHD_ETABLISSEMENT="" AND PHD_SALARIE="" AND PHD_CODEPOP="" AND '
  else WhereInfo := '';
  GrilleEC := THGrid(GetControl('GELTCOMPL'));
  Ligne := GrilleEC.Row;
  Tablette := '';
  LeType:= GrilleEC.CellValues[5,Ligne];
  If LeType = 'THEME' then exit;
  Champ := GrilleEC.CellValues[6,Ligne];
  Libelle := GrilleEC.CellValues[1,Ligne];
  If typeSaisie ='POP' then Code := CodePop
  else if typeSaisie = 'ETB' then Code := Etabliss
  else Code := '';
  If LeType = 'B' then Ret := AGLLanceFiche('PAY','PGLISTEHISTODH','','',TypeSaisie+';'+Code+';'+Champ+';;'+TypePop)
  else If (LeType = 'C') or (LeType = 'T') then Ret := AGLLanceFiche('PAY','PGLISTEHITODCL','','',TypeSaisie+';'+Code+';'+Champ+';'+'PGCOMBOZONELIBRE;'+TypePop)
//PT1  else Ret := AGLLanceFiche('PAY','PGLISTEHISTODE','','',TypeSaisie+';'+Salarie+';'+Code+';'+Tablette+';'+TypePop);
  else Ret := AGLLanceFiche('PAY','PGLISTEHISTODE','','',TypeSaisie+';'+Code+';'+Champ+';'+Tablette+';'+TypePop); //PT1
  If Ret <> '' then
  begin
    Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE '+WhereInfo+
      'PHD_PGINFOSMODIF="'+GrilleEC.CellValues[6,Ligne]+'" AND PHD_DATEAPPLIC<="'+UsDateTime(V_PGI.DateEntree)+'" ORDER BY PHD_DATEAPPLIC DESC',True);
    Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
    DateApplic := Q.FindField('PHD_DATEAPPLIC').AsDateTime;
    Ferme(Q);
    Niveau := GrilleEC.CellValues[0,Ligne];
    GrilleEC.CellValues[2,Ligne] := Valeur;
    GrilleEC.CellValues[3,Ligne] := DateToStr(DateApplic);
    GrilleEC.CellValues[4,Ligne] := Img;
  end;
end;

procedure TOF_PGSAISIEZL.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var GrilleEC : THGrid;
begin
GrilleEC := THGrid(GetControl('GELTCOMPL'));
 If (GrilleEC.CellValues[5,ARow] = 'THEME') and Not(GrilleEC.IsSelected(ARow)) then
  begin
//    If GrilleEC.Row = ARow then Canvas.Font.Color := ClBlue;
    Canvas.Font.Style := [fsBold];
    //Canvas.Brush.Color := ClBtnFace;
    Canvas.Brush.Color := ClInactiveCaption;
    end;
 end;


Initialization
  registerclasses ( [ TOF_PGSAISIEZL ] ) ;
end.

