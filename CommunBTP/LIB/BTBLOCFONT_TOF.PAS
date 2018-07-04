{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/03/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTBLOCFONT ()
Mots clefs ... : TOF;BTBLOCFONT
*****************************************************************}
Unit BTBLOCFONT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,fe_main,
{$else}
     eMul,eagl,
{$ENDIF}
		 UtilPgi,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     vierge,
     graphics,
     dialogs,
     HTB97,
     AglInit,
     Hrichole,
     UTOF ;

Type
  TOF_BTBLOCFONT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fToolWinSrc : TToolWindow97;
    FtexteSrc : THRichEditOle;
    fCurrFont : Tfont;
    fChampApplication : string;
    fTOB : TOB;
    fTOBTableliste : TOB;
    procedure DecodeArg (S : string);
    procedure FontElipsisClick (Sender : Tobject);
    procedure ApplicfontClick (Sender : Tobject);
    procedure ConstitueListeTables;
    procedure AjouteTable(NomTable, Libelle: string);
    procedure Appliquelapolice;
    procedure TraitelaTable(TOBL: TOB);
    procedure UpdateDecoupeAffaire;
    procedure UpdateDecoupeArticle;
    procedure UpdateDecoupeCommercial;
    procedure UpdateDecoupeContact;
    procedure UpdateDecoupeLiensOle;
    procedure UpdateDecoupeLigne;
    procedure UpdateDecoupePiece;
    procedure UpdateDecoupeSuspect;
    procedure UpdateDecoupeTiers;
    procedure SetWidth(largeur: integer);
  end ;

procedure ParamFont(ChampsApplication : string);

Implementation
uses BTPUtil,FactUtil,UtilTOBPiece,uEntCommun,FactTOB ;

procedure ParamFont(ChampsApplication : string);
begin
  AGLLanceFiche('BTP','BTBLOCFONT','','','ACTION=MODIFICATION;APPLICATION='+ChampsApplication);
end;


procedure TOF_BTBLOCFONT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCFONT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCFONT.OnUpdate ;
begin
  fTOB.GetEcran(ecran);
  if fTOB.GetString('BFS_FONT')='' then
  begin
    if (PGIAsk('Etes-vous sur de ne définir aucune police de caractères ?') <> mryes) then
    begin
      TFVierge(Ecran).ModalResult := 0;
      exit;
    end;
  end;
  fTOB.InsertOrUpdateDB(false); 
  Inherited ;
end ;

procedure TOF_BTBLOCFONT.OnLoad ;
var SQL : string;
		QQ : TQuery;
begin
  Inherited ;
	Sql := 'SELECT * FROM BFONTS WHERE BFS_APPLICATION="'+fChampApplication+'"';
  QQ := OpenSql (Sql,true,1,'',true);
  if not QQ.eof then
  begin
		fTOB.SelectDB('',QQ);
  end else
  begin
    fTOB.Setstring('BFS_APPLICATION',fChampApplication);
    fTOB.SetString('BFS_FONT',V_PGI.FontMemo);
  end;
  fTOB.PutEcran(ecran); 
  ferme (QQ);
  DecodeFont (THEdit(GetControl('BFS_FONT')).Text,fCurrFont);
end ;

procedure TOF_BTBLOCFONT.OnArgument (S : String ) ;
begin
  Inherited ;
  fCurrFont := Tfont.Create;
  fTOB := TOB.Create ('BFONTS',nil,-1);
  fTOBTableliste := TOB.Create ('LES TABLES',nil,-1);
  fTOBTableliste.AddChampSupValeur ('OKOK','-');
  //
  DecodeArg (S);
  THEdit(GetControl('BFS_FONT')).OnElipsisClick := FontElipsisClick;
  TToolbarButton97 (GetControl('BAPPLICFONT')).OnClick := ApplicfontClick;
  //
  fToolWinSrc := TToolWindow97.Create(TFvierge(Ecran));
  fToolWinSrc.parent := ecran;
  FtexteSrc := THRichEditOle.Create(fToolWinSrc);
  FtexteSrc.Parent := fToolWinSrc;
  FtexteSrc.Align := alClient;
  fToolWinSrc.Width := 500;
  fToolWinSrc.Visible  := false;
  //
  ConstitueListeTables;

end ;

procedure TOF_BTBLOCFONT.OnClose ;
begin
  fToolWinSrc.free;
  fTOBTableliste.free;
  fTOB.free;
  Inherited ;
end ;

procedure TOF_BTBLOCFONT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCFONT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCFONT.DecodeArg(S: string);
var LesCriteres,critere : string;
		Champ,Valeur : string;
    X : integer;
begin
  //Récupération valeur de argument
  LesCriteres := S;
  repeat
    Critere:=(Trim(ReadTokenSt(LesCriteres)));
    if (Critere <> '') then
    begin
      X := pos ('=', Critere) ;
      if x <> 0 then
      begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        if Champ = 'APPLICATION' then
        begin
          fChampApplication := valeur;
        end;
      end
    end;
  until critere = '';
end;

procedure TOF_BTBLOCFONT.SetWidth (largeur : integer);
begin
	fToolWinSrc.Width := largeur;
  fToolWinSrc.Refresh;
  FtexteSrc.Clear;
end;

procedure TOF_BTBLOCFONT.FontElipsisClick(Sender: Tobject);
var TheFontDialog : TFontDialog ;
begin
	TheFontDialog := TFontDialog.Create(TFVierge(Ecran));
  TRY
    DecodeFont (fTOB.GetString('BFS_FONT'), TheFontDialog.Font);
    TheFontDialog.Options := TheFontDialog.Options- [fdEffects]; 
    if TheFontDialog.Execute then
    begin
      fTOB.SetString('BFS_FONT',EncodeFont(TheFontDialog.font) );
      fTOB.PutEcran(ecran); 
  		DecodeFont (THEdit(GetControl('BFS_FONT')).Text,fCurrFont);
    end;
  FINALLY
  	TheFontDialog.free;
  END;
end;

procedure TOF_BTBLOCFONT.Appliquelapolice;
var Indice : integer;
begin
	For Indice := 0 to fTOBTableliste.detail.count -1 do
  begin
    if fTOBTableliste.detail[Indice].getString('SELECT')='X' then
    begin
			TraitelaTable (fTOBTableliste.detail[Indice]);
    end;
  end;
end;

procedure TOF_BTBLOCFONT.UpdateDecoupeArticle;
var lastArt,Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  fTOBA := TOB.create ('LES ARTICLES',nil,-1);
  lastArt := '';
  SetWidth (600); // largeur depuis la fiche Article
	repeat
    fTOBA.ClearDetail;
    Sql := 'SELECT ##TOP 100## GA_ARTICLE,GA_BLOCNOTE FROM ARTICLE';
    if lastArt <> '' then Sql := Sql + ' WHERE GA_ARTICLE > "'+lastArt+'"';
    Sql := Sql + ' ORDER BY GA_ARTICLE';
    QQ := OpenSql (Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      fTOBA.LoadDetailDB('ARTICLE','','',QQ,false);
      LastArt := ftobA.detail[fTOBA.detail.count-1].GetString('GA_ARTICLE');
    end;
    ferme (QQ);
    if FTOBA.Detail.count > 0 then
    begin
      for Indice := 0 to FTOBA.Detail.count -1 do
      begin
      	StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('GA_BLOCNOTE'));
        if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
        begin
					// Si le bloc note est vide on ne traite pas
          FtexteSrc.SelectAll;
          FtexteSrc.SelAttributes.Name := fCurrFont.Name;
          FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
          FtexteSrc.SelAttributes.Size   := fCurrFont.size;
          FTOBA.Detail[Indice].PutValue('GA_BLOCNOTE',ExRichToString(FtexteSrc));
          fTOBA.Detail[Indice].UpdateDB(false)
        end;
      end;
    end;
  until fTOBA.detail.count = 0;
  fTOBA.free;
end;

procedure TOF_BTBLOCFONT.UpdateDecoupeSuspect;
var lastArt,Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  fTOBA := TOB.create ('LES SUSPECTS',nil,-1);
  lastArt := '';
  SetWidth (600); // largeur depuis la fiche Article
	repeat
    fTOBA.ClearDetail;
    Sql := 'SELECT ##TOP 100## RSU_SUSPECT,RSU_BLOCNOTE FROM SUSPECTS';
    if lastArt <> '' then Sql := Sql + ' WHERE RSU_SUSPECT > "'+lastArt+'"';
    Sql := Sql + ' ORDER BY RSU_SUSPECT';
    QQ := OpenSql (Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      fTOBA.LoadDetailDB('SUSPECTS','','',QQ,false);
      LastArt := ftobA.detail[fTOBA.detail.count-1].GetString('RSU_SUSPECT');
    end;
    ferme (QQ);
    if FTOBA.Detail.count > 0 then
    begin
      for Indice := 0 to FTOBA.Detail.count -1 do
      begin
      	StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('RSU_BLOCNOTE'));
        if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
        begin
					// Si le bloc note est vide on ne traite pas
          FtexteSrc.SelectAll;
          FtexteSrc.SelAttributes.Name := fCurrFont.Name;
          FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
          FtexteSrc.SelAttributes.Size   := fCurrFont.size;
          FTOBA.Detail[Indice].PutValue('RSU_BLOCNOTE',ExRichToString(FtexteSrc));
          fTOBA.Detail[Indice].UpdateDB(false)
        end;
      end;
    end;
  until fTOBA.detail.count = 0;
  fTOBA.free;

end;

procedure TOF_BTBLOCFONT.UpdateDecoupeTiers;
var lastArt,Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  fTOBA := TOB.create ('LES TIERS',nil,-1);
  lastArt := '';
  SetWidth (600); // largeur depuis la fiche Article
	repeat
    fTOBA.ClearDetail;
    Sql := 'SELECT ##TOP 100## T_AUXILIAIRE,T_BLOCNOTE FROM TIERS';
    if lastArt <> '' then Sql := Sql + ' WHERE T_AUXILIAIRE > "'+lastArt+'"';
    Sql := Sql + ' ORDER BY T_AUXILIAIRE';
    QQ := OpenSql (Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      fTOBA.LoadDetailDB('TIERS','','',QQ,false);
      LastArt := ftobA.detail[fTOBA.detail.count-1].GetString('T_AUXILIAIRE');
    end;
    ferme (QQ);
    if FTOBA.Detail.count > 0 then
    begin
      for Indice := 0 to FTOBA.Detail.count -1 do
      begin
      	StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('T_BLOCNOTE'));
        if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
        begin
					// Si le bloc note est vide on ne traite pas
          FtexteSrc.SelectAll;
          FtexteSrc.SelAttributes.Name := fCurrFont.Name;
          FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
          FtexteSrc.SelAttributes.Size   := fCurrFont.size;
          FTOBA.Detail[Indice].PutValue('T_BLOCNOTE',ExRichToString(FtexteSrc));
          TRY
          	fTOBA.Detail[Indice].UpdateDB(false)
          EXCEPT
            continue;
          END;
        end;
      end;
    end;
  until fTOBA.detail.count = 0;
  fTOBA.free;

end;

procedure TOF_BTBLOCFONT.UpdateDecoupeContact;
var lastclef : string;
		Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  lastclef := '';
  fTOBA := TOB.create ('LES CONTACTS',nil,-1);
  TRY
    lastclef  := '';
    SetWidth (600); // largeur depuis la fiche Article
    repeat
      fTOBA.ClearDetail;
      Sql := 'SELECT ##TOP 100## (C_TYPECONTACT+C_AUXILIAIRE+CONVERT(VARCHAR(10),C_NUMEROCONTACT)) AS CLEF,'+
      			 'C_TYPECONTACT,C_AUXILIAIRE,C_NUMEROCONTACT,C_BLOCNOTE FROM CONTACT';
      if lastClef <> '' then
      begin
      	Sql := Sql + ' WHERE C_TYPECONTACT+C_AUXILIAIRE+CONVERT(VARCHAR(10),C_NUMEROCONTACT) > "'+LastClef+'"' ;
      END;
      Sql := Sql + ' ORDER BY CLEF';
      QQ := OpenSql (Sql,true,-1,'',true);
      if not QQ.eof then
      begin
        fTOBA.LoadDetailDB('CONTACT','','',QQ,false);
        lastClef := ftobA.detail[fTOBA.detail.count-1].GetString('CLEF');
      end;
      ferme (QQ);
      if FTOBA.Detail.count > 0 then
      begin
        for Indice := 0 to FTOBA.Detail.count -1 do
        begin
          StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('C_BLOCNOTE'));
          if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
          begin
            // Si le bloc note est vide on ne traite pas
            FtexteSrc.SelectAll;
            FtexteSrc.SelAttributes.Name := fCurrFont.Name;
            FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
            FtexteSrc.SelAttributes.Size   := fCurrFont.size;
            FTOBA.Detail[Indice].PutValue('C_BLOCNOTE',ExRichToString(FtexteSrc));
            fTOBA.Detail[Indice].UpdateDB(false)
          end;
        end;
      end;
    until fTOBA.detail.count = 0;
  FINALLY
  	fTOBA.free;
  END;

end;

procedure TOF_BTBLOCFONT.UpdateDecoupeCommercial;
var lastArt,Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  fTOBA := TOB.create ('LES COMMERCIAUX',nil,-1);
  lastArt := '';
  SetWidth (600); // largeur depuis la fiche Article
	repeat
    fTOBA.ClearDetail;
    Sql := 'SELECT ##TOP 100## GCL_COMMERCIAL,GCL_BLOCNOTE FROM COMMERCIAL';
    if lastArt <> '' then Sql := Sql + ' WHERE GCL_COMMERCIAL > "'+lastArt+'"';
    Sql := Sql + ' ORDER BY GCL_COMMERCIAL';
    QQ := OpenSql (Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      fTOBA.LoadDetailDB('COMMERCIAL','','',QQ,false);
      LastArt := ftobA.detail[fTOBA.detail.count-1].GetString('GCL_COMMERCIAL');
    end;
    ferme (QQ);
    if FTOBA.Detail.count > 0 then
    begin
      for Indice := 0 to FTOBA.Detail.count -1 do
      begin
      	StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('GCL_BLOCNOTE'));
        if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
        begin
					// Si le bloc note est vide on ne traite pas
          FtexteSrc.SelectAll;
          FtexteSrc.SelAttributes.Name := fCurrFont.Name;
          FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
          FtexteSrc.SelAttributes.Size   := fCurrFont.size;
          FTOBA.Detail[Indice].PutValue('GCL_BLOCNOTE',ExRichToString(FtexteSrc));
          fTOBA.Detail[Indice].UpdateDB(false)
        end;
      end;
    end;
  until fTOBA.detail.count = 0;
  fTOBA.free;

end;

procedure TOF_BTBLOCFONT.UpdateDecoupeAffaire;
var lastArt,Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  fTOBA := TOB.create ('LES AFFAIRES',nil,-1);
  lastArt := '';
  SetWidth (600); // largeur depuis la fiche Article
	repeat
    fTOBA.ClearDetail;
    Sql := 'SELECT ##TOP 100## AFF_AFFAIRE,AFF_DESCRIPTIF FROM AFFAIRE';
    if lastArt <> '' then Sql := Sql + ' WHERE AFF_AFFAIRE > "'+lastArt+'"';
    Sql := Sql + ' ORDER BY AFF_AFFAIRE';
    QQ := OpenSql (Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      fTOBA.LoadDetailDB('AFFAIRE','','',QQ,false);
      LastArt := ftobA.detail[fTOBA.detail.count-1].GetString('AFF_AFFAIRE');
    end;
    ferme (QQ);
    if FTOBA.Detail.count > 0 then
    begin
      for Indice := 0 to FTOBA.Detail.count -1 do
      begin
      	StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('AFF_DESCRIPTIF'));
        if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
        begin
					// Si le bloc note est vide on ne traite pas
          FtexteSrc.SelectAll;
          FtexteSrc.SelAttributes.Name := fCurrFont.Name;
          FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
          FtexteSrc.SelAttributes.Size   := fCurrFont.size;
          FTOBA.Detail[Indice].PutValue('AFF_DESCRIPTIF',ExRichToString(FtexteSrc));
          fTOBA.Detail[Indice].UpdateDB(false)
        end;
      end;
    end;
  until fTOBA.detail.count = 0;
  fTOBA.free;
end;

procedure TOF_BTBLOCFONT.UpdateDecoupePiece;
var lastclef : string;
		Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  lastclef := '';
  fTOBA := TOB.create ('LES PIECES',nil,-1);
  TRY
    lastclef  := '';
    SetWidth (600); // largeur depuis la fiche Article
    repeat
      fTOBA.ClearDetail;
      Sql := 'SELECT ##TOP 100## (GP_NATUREPIECEG+GP_SOUCHE+CONVERT(VARCHAR(5),GP_NUMERO)+CONVERT(VARCHAR(2),GP_INDICEG)) AS CLEF,'+
      			 'GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_BLOCNOTE FROM PIECE';
      if lastClef <> '' then
      begin
      	Sql := Sql + ' WHERE GP_NATUREPIECEG+GP_SOUCHE+CONVERT(VARCHAR(5),GP_NUMERO)+CONVERT(VARCHAR(2),GP_INDICEG) > "'+LastClef+'"' ;
      END;
      Sql := Sql + ' ORDER BY CLEF';
      QQ := OpenSql (Sql,true,-1,'',true);
      if not QQ.eof then
      begin
        fTOBA.LoadDetailDB('PIECE','','',QQ,false);
        lastClef := ftobA.detail[fTOBA.detail.count-1].GetString('CLEF');
      end;
      ferme (QQ);
      if FTOBA.Detail.count > 0 then
      begin
        for Indice := 0 to FTOBA.Detail.count -1 do
        begin
          StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('GP_BLOCNOTE'));
          if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
          begin
            // Si le bloc note est vide on ne traite pas
            FtexteSrc.SelectAll;
            FtexteSrc.SelAttributes.Name := fCurrFont.Name;
            FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
            FtexteSrc.SelAttributes.Size   := fCurrFont.size;
            FTOBA.Detail[Indice].PutValue('GP_BLOCNOTE',ExRichToString(FtexteSrc));
            fTOBA.Detail[Indice].UpdateDB(false)
          end;
        end;
      end;
    until fTOBA.detail.count = 0;
  FINALLY
  	fTOBA.free;
  END;
end;

procedure TOF_BTBLOCFONT.UpdateDecoupeLigne;
var lastclef : string;
		Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
    cledoc : r_cledoc;
    SS : WideString;
begin
  lastclef := '';
  fTOBA := TOB.create ('LES LIGNES',nil,-1);
  TRY
    lastclef  := '';
    SetWidth (600); // largeur depuis la fiche Article
    repeat
      fTOBA.ClearDetail;
      Sql := 'SELECT ##TOP 100## (GL_NATUREPIECEG+GL_SOUCHE+CONVERT(VARCHAR(10),GL_NUMERO)+CONVERT(VARCHAR(2),GL_INDICEG)+CONVERT(VARCHAR(7),GL_NUMORDRE)) AS CLEF,'+
      			 'GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMORDRE,GL_BLOCNOTE FROM LIGNE WHERE NOT GL_BLOCNOTE IS NULL';  // WHERE GL_NATUREPIECEG="DBT" AND GL_NUMERO=2013090';
      if lastClef <> '' then
      begin
      	Sql := Sql + ' AND GL_NATUREPIECEG+GL_SOUCHE+CONVERT(VARCHAR(10),GL_NUMERO)+CONVERT(VARCHAR(2),GL_INDICEG)+CONVERT(VARCHAR(7),GL_NUMORDRE) > "'+LastClef+'"' ;
      END;
      Sql := Sql + ' ORDER BY CLEF';
      QQ := OpenSql (Sql,true,-1,'',true);
      if not QQ.eof then
      begin
        fTOBA.LoadDetailDB('LIGNE','','',QQ,false);
        lastClef := ftobA.detail[fTOBA.detail.count-1].GetString('CLEF');
      end;
      ferme (QQ);
      if FTOBA.Detail.count > 0 then
      begin
        for Indice := 0 to FTOBA.Detail.count -1 do
        begin
          StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('GL_BLOCNOTE'));
          if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
          begin
            Cledoc := TOB2Cledoc(FTOBA.Detail[Indice]);
            // Si le bloc note est vide on ne traite pas
            FtexteSrc.SelectAll;
            FtexteSrc.SelAttributes.Name := fCurrFont.Name;
            FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
            FtexteSrc.SelAttributes.Size   := fCurrFont.size;
            (*
            FTOBA.Detail[Indice].PutValue('GL_BLOCNOTE',RichToString(FtexteSrc));
            FTOBA.Detail[Indice].SetAllModifie(true);
            *)
            SS := ExRichToString(FtexteSrc);
            SS := StringReplace(SS,'"','""',[rfreplaceall]);
            SQL := 'UPDATE LIGNE SET GL_BLOCNOTE="'+SS+'" WHERE ' +
            						WherePiece(cledoc,ttdLigne,True,true);
            ExecuteSQL(SQL);
//            fTOBA.Detail[Indice].UpdateDB(false)
          end;
        end;
      end;
    until fTOBA.detail.count = 0;
  FINALLY
  	fTOBA.free;
  END;
end;

procedure TOF_BTBLOCFONT.UpdateDecoupeLiensOle;
var lastclef : string;
		Sql  : string;
		QQ : TQuery;
    fTOBA : TOB;
    Indice : integer;
begin
  lastclef := '';
  fTOBA := TOB.create ('LES LIENSOLE',nil,-1);
  TRY
    lastclef  := '';
    SetWidth (600); // largeur depuis la fiche Article
    repeat
      fTOBA.ClearDetail;
      Sql := 'SELECT ##TOP 100## (LO_TABLEBLOB+LO_IDENTIFIANT+CONVERT(VARCHAR(4),LO_RANGBLOB)) AS CLEF,'+
      			 'LO_TABLEBLOB,LO_IDENTIFIANT,LO_RANGBLOB,LO_OBJET FROM LIENSOLE';
      if lastClef <> '' then
      begin
      	Sql := Sql + ' WHERE LO_TABLEBLOB+LO_IDENTIFIANT+CONVERT(VARCHAR(4),LO_RANGBLOB) > "'+LastClef+'"' ;
      END;
      Sql := Sql + ' ORDER BY CLEF';
      QQ := OpenSql (Sql,true,-1,'',true);
      if not QQ.eof then
      begin
        fTOBA.LoadDetailDB('LIENSOLE','','',QQ,false);
        lastClef := ftobA.detail[fTOBA.detail.count-1].GetString('CLEF');
      end;
      ferme (QQ);
      if FTOBA.Detail.count > 0 then
      begin
        for Indice := 0 to FTOBA.Detail.count -1 do
        begin
          StringToRich(FtexteSrc, FTOBA.Detail[Indice].GetValue ('LO_OBJET'));
          if (Length(FtexteSrc.Text) <> 0) and (FtexteSrc.Text <> #$D#$A) then
          begin
            // Si le bloc note est vide on ne traite pas
            FtexteSrc.SelectAll;
            FtexteSrc.SelAttributes.Name := fCurrFont.Name;
            FtexteSrc.SelAttributes.Style  := fCurrFont.Style;
            FtexteSrc.SelAttributes.Size   := fCurrFont.size;
            FTOBA.Detail[Indice].PutValue('LO_OBJET',ExRichToString(FtexteSrc));
            fTOBA.Detail[Indice].UpdateDB(false)
          end;
        end;
      end;
    until fTOBA.detail.count = 0;
  FINALLY
  	fTOBA.free;
  END;
end;

procedure TOF_BTBLOCFONT.TraitelaTable (TOBL : TOB);
begin
	if TOBL.GetString ('TABLE')='ARTICLE' then
  begin
    UpdateDecoupeArticle;
  end else if TOBL.GetString ('TABLE')='SUSPECTS' then
  begin
    UpdateDecoupeSuspect;
  end else if TOBL.GetString ('TABLE')='TIERS' then
  begin
    UpdateDecoupeTiers;
  end else if TOBL.GetString ('TABLE')='CONTACT' then
  begin
    UpdateDecoupeContact;
  end else if TOBL.GetString ('TABLE')='COMMERCIAL' then
  begin
    UpdateDecoupeCommercial;
  end else if TOBL.GetString ('TABLE')='AFFAIRE' then
  begin
    UpdateDecoupeAffaire;
  end else if TOBL.GetString ('TABLE')='PIECE' then
  begin
    UpdateDecoupePiece;
  end else if TOBL.GetString ('TABLE')='LIGNE' then
  begin
    UpdateDecoupeLigne;
  end else if TOBL.GetString ('TABLE')='LIENSOLE' then
  begin
    UpdateDecoupeLiensOle;
  end;
end;

procedure TOF_BTBLOCFONT.ApplicfontClick(Sender: Tobject);
var Indice : integer;
begin
  fTOBTableliste.SetString('OKOK','-');
	For Indice := 0 to fTOBTableliste.detail.count -1 do
  begin
  	fTOBTableliste.detail[indice].SetString('SELECT','-');
  end;

  TheTOB := fTOBTableliste;
  TRY
    AGLLanceFiche('BTP','BTBLOCSELECT','','','');
    if fTOBTableliste.GetString('OKOK') = 'X' then
    begin
      Appliquelapolice;
    end;
  FINALLY
    TheTOB := nil;
  END;
end;

procedure TOF_BTBLOCFONT.AjouteTable (NomTable : string; Libelle : string) ;
var TOBL: TOB;
begin
	TOBL := TOB.Create ('UNE TABLE',fTOBTableliste,-1);
  TOBL.AddChampSupValeur('OPEN','');
  TOBL.AddChampSupValeur('TABLE',Nomtable);
  TOBL.AddChampSupValeur('LIBELLE',Libelle);
  TOBL.AddChampSupValeur('SELECT','-');
end;

procedure TOF_BTBLOCFONT.ConstitueListeTables;
begin
  AjouteTable ('ARTICLE','Table Article');
  AjouteTable ('SUSPECTS','Table des Suspects');
  AjouteTable ('TIERS','Table des Tiers');
  AjouteTable ('CONTACT','Table Contacts des Tiers');
  AjouteTable ('COMMERCIAL','Table des commerciaux');
  AjouteTable ('AFFAIRE','Table des Affaires/Contrats/Interventions');
  AjouteTable ('PIECE','Descriptif entête de document');
  AjouteTable ('LIGNE','Descriptif lignes de document');
  AjouteTable ('LIENSOLE','Reste');
end;

Initialization
  registerclasses ( [ TOF_BTBLOCFONT ] ) ;
end.
