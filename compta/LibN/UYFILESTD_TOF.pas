unit UYFILESTD_TOF;

interface

Uses StdCtrls, Controls, Classes, Windows,
{$IFDEF EAGLCLIENT}
     eFiche,eFichList, Maineagl,
{$ELSE}
     db,{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}Fiche, FichList,  Fe_Main,
{$ENDIF}
     forms, sysutils,ComCtrls,HCtrls, HEnt1, HMsgBox, UTOF, dialogs,
{$IFDEF SCANGED}
    YNewDocument, UGedFiles, AnnOutils, uTilGed,
{$ENDIF SCANGED}
     UTob, HPanel, uYFILESTD, HTB97, Lookup, Ent1, ULibWindows, Vierge, HRichOle, Cbppath;


Type
  TOF_YFILESTDXL = Class (TOF)
   procedure OnArgument(stArgument: String); override ;
   procedure OnUpdate                      ; override ;
   procedure OnNew                         ; override ;
   procedure OnLoad                        ; override;
   procedure OnClose ; override ;
   private
   SType, FichierAIntegrer                 : string;
   LeMode                                  : TActionFiche ;
   Pages                                   : TPageControl;
   SaisieCommentaire                       : Boolean;
   FTobFts                                 : TOB;
   DocAintegrer                            : string;
   OkRef                                   : Boolean;
   procedure BIBLIOnClick(Sender: TObject);
   procedure BREFOnClick(Sender: TObject);
   procedure DupliqueOnClick(Sender: TObject);

end;
procedure CPLanceFiche_YFILESTDXL (Cle, TypS : string);
{$IFDEF SCANGED}
Procedure MyImportGed (FileGuid: string; Lib : string);
{$ENDIF}

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  Paramsoc;


procedure CPLanceFiche_YFILESTDXL (Cle, TypS : string);
begin
    AglLanceFiche('CP','CPIMPFICHFTS','', Cle, TypS)
end;

procedure TOF_YFILESTDXL.OnArgument(stArgument: String);
var
St : string;
GBCarac, GBCrit : TGroupBox;
begin
  Inherited ;
        if stArgument <> '' then
        begin
           St := StArgument;
           SType := ReadTokenSt(St);
           SaisieCommentaire := FALSE;
           if St= 'ACTION=CREATION' then
              LeMode:=taCreat
           else
           begin
                LeMode := taModif;
                if St = 'ACTION=COMMENTAIRE' then
                   SaisieCommentaire := TRUE;
           end;
        end;
        if (SType <> '') then
        begin
              SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
              SetControlText ('YFS_CRIT1', SType);
              SetControlEnabled ('YFS_CODEPRODUIT', FALSE);
              if LeMode=taCreat then
              begin
                 TToolBarButton97( GetControl('BIBLI') ).OnClick          := BIBLIOnClick;
                 TToolBarButton97(GetControl('BREF')).OnClick             := BREFOnClick;
              end;
              TToolBarButton97(GetControl('DUPLIQUER')).OnClick        := DupliqueOnClick;
        end;
        if LeMode = taCreat then
          SetControlvisible ('DUPLIQUER', FALSE);
        FichierAIntegrer := '';
       FTobFts := TFVierge(Ecran).laTof.laTob ;
       FTobFts.PutEcran( Ecran ) ;
       if SaisieCommentaire then
       begin
          SetControlvisible ('DUPLIQUER', FALSE);
          SetControlvisible ('BREF', FALSE);
          SetControlvisible ('BIBLI', FALSE);
       end;
      DocAintegrer := '';
   if SType = 'EXL' then
          Ecran.Caption := 'Autres feuilles Excel'
   else
   if SType = 'FTS' then
          Ecran.Caption := 'Feuilles de travail spécialisées';

  // Affichage des GroupBox
  if SType = 'EXL' then
  begin
    GBCarac := TGroupBox(GetControl('GBCARACT'));
    GBCrit := TGroupBox(GetControl('GPEBCYCLE'));
    if GBCarac <> nil then
    begin
      GBCarac.Visible := False;
      if GBCrit <> nil then
      begin
        GBCrit.Left := GBCarac.Left;
        GBCrit.Top := GBCarac.Top;
      end;
    end;
  end;

  UpdateCaption( Ecran ) ;

end;

procedure TOF_YFILESTDXL.OnLoad ;
var
Nom       : string;
lStVal    : string;
begin
  Inherited ;
        Nom := THEdit (GetControl ('YFS_NOM')).Text;
        Pages := TPageControl(GetControl('PAGES', true));

        SetControlText ('FICHENAME', Nom);
       if (SType <> '') then
       begin
             if(LeMode in [taCreat..taCreatOne]) then
              begin
                     OnNew;
                     if EstSpecif('51502') then
                        lStVal := 'CEG'
                     else
                     if (ctxStandard in V_PGI.PGIContexte) then
                     begin
                        lStVal := 'STD';
                        V_PGI.NoDossier := '000000';
                     end
                     else
                        lStVal := 'DOS';

                     SetControlText ('YFS_PREDEFINI', lStVal);
              end
              else
              if (LeMode = taModif) then
              begin
                      if not EstSpecif('51502') then
                      begin
                           if (THValComboBox (GetControl ('YFS_PREDEFINI')).value = 'CEG') then
                           begin
                               SetControlEnabled ('PGENERAL', FALSE);
                               Ecran.Caption := THEdit (GetControl ('YFS_NOM')).Text + ' : Standard CEGID';
                           end;
                           if not (ctxStandard in V_PGI.PGIContexte) and (THValComboBox (GetControl ('YFS_PREDEFINI')).value = 'STD') then
                           begin
                                     SetControlEnabled ('PGENERAL', FALSE);
                                     Ecran.Caption := THEdit (GetControl ('YFS_NOM')).Text + ' : Standard CABINET';
                           end
                      end;
                      SetControlEnabled ('FICHENAME', FALSE);
              end;
            SetControlEnabled ('YFS_PREDEFINI', FALSE);
            if SaisieCommentaire then
            begin
               Pages.Pages[0].TabVisible := FALSE;
               Pages.Pages[1].TabVisible := TRUE;
            end
            else
            begin
               SetControlEnabled ('YFS_CRIT1', FALSE);
               SetControlEnabled ('YFS_CRIT2', TRUE);
               if  lStVal <> 'STD' then
               begin
                  THEdit(GetControl ('YFS_CRIT3')).Datatype := 'CREVCYCLEACTIF';
                  // fiche 21182
                  THEdit(GetControl ('YFS_CRIT3')).Plus := 'AND CCY_EXERCICE="'+GetEncours.Code+'"';
               end
               else
                  THEdit(GetControl ('YFS_CRIT3')).Datatype := 'CREVPARAMCYCLE';
               if not (ctxStandard in V_PGI.PGIContexte) then
                  SetControlEnabled ('YFS_CRIT3', (GetParamSocSecur('SO_CPPLANREVISION', '') <> ''));
               SetControlEnabled ('YFS_CRIT4', TRUE);
               SetControlEnabled ('YFS_CRIT5', TRUE);
            end;

            lStVal := FormatDateTime( 'dd mmm yyyy', FTobFts.getvalue ('YFS_DATEMODIF') ) ;
            SetControlText('DATEMODIF', lStVal ) ;
            lStVal := FormatDateTime( 'dd mmm yyyy', FTobFts.getvalue ('YFS_DATECREATION') ) ;
            SetControlText('DATECREATION', lStVal ) ;
           THEdit(GetControl('YFS_CRIT4')).Plus := ' AND YFS_CRIT1="'+SType+'"';
           THEDIT(GetControl('YFS_CRIT5')).Plus := ' AND YFS_CRIT1="'+SType+'"';
      end;
end;


procedure TOF_YFILESTDXL.OnUpdate;
var
Numdos, SQL   : string;
Ext,Fname     : string;
Filename      : string;
Ind, i        : integer;
ExisteS       : Boolean;
begin
  Inherited ;
        SetControlText ('YFS_NOM', UpperCase(ExtractFileName(THEdit (GetControl ('FICHENAME')).Text)));
        if Length (THEdit (GetControl ('YFS_NOM')).Text) > 70 then
        begin
             PgiInfo ('Attention , la longueur du fichier ne doit pas dépasser de 70 caractères');  LastError := -1; exit;
        end;

        if EstSpecif('51502') then
                Numdos := '000000'
        else
                Numdos := V_PGI.NoDossier;

        if ((SType <> '') and (LeMode = taCreat)) or (SType = '')  then
        begin
            SQL := 'SELECT YFS_FILEGUID FROM YFILESTD Where YFS_CODEPRODUIT="COMPTA" AND '+
                ' YFS_CRIT1="'+THEdit (GetControl ('YFS_CRIT1')).Text+'" AND'+
                ' YFS_NOM="'+THEdit (GetControl ('YFS_NOM')).Text+'" AND'+
                ' YFS_LANGUE="'+V_PGI.LanguePrinc+'" AND'+
                ' YFS_NODOSSIER="'+Numdos+'" AND'+
                ' YFS_PREDEFINI="'+THValComboBox (GetControl ('YFS_PREDEFINI')).Value+'"';
             ExisteS :=ExisteSQL (SQL);
            if ExisteS then
             begin
                 Ext := THEdit (GetControl ('YFS_NOM')).Text;
                 Fname   := ReadTokenPipe (Ext, '.');
                 Ext := '.'+Ext;
                 Ind := 1;
                 i :=  Pos ('_', Fname);
                 if i > 0 then
                 Fname := copy(Fname, 0, i-1);

                 While (ExisteSql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+Fname+ IntTostr(Ind)+Ext+'" AND YFS_CRIT1="'+SType+
                 '" AND YFS_PREDEFINI="'+THValComboBox(GetControl('YFS_PREDEFINI')).Value+'" AND YFS_NODOSSIER="'+ V_PGI.NoDossier+'"')) do inc (Ind);
                 Fname   := Fname+ IntTostr(Ind)+Ext;
                 FichierAIntegrer := Fname;

               if OkRef then
               begin
                      FichierAIntegrer := ExtractFileDir(DocAintegrer)+ '\'+ Fname;
                      Copyfile (PChar( DocAintegrer), Pchar(FichierAIntegrer),TRUE);
               end
               else
               begin
                      FichierAIntegrer := ExtractFileDir(THEdit (GetControl ('FICHENAME')).Text)+'\'+ Fname;
                      Filename := THEdit (GetControl ('FICHENAME')).Text;
                      Copyfile (PChar( Filename), Pchar(FichierAIntegrer),TRUE);
                end;
                SetControlText ('YFS_NOM', ExtractFileName(FichierAIntegrer));
             end
             else
             begin
                if OkRef then  FichierAIntegrer := DocAintegrer     // fiche  20862
                else
                if FichierAIntegrer = '' then   FichierAIntegrer := THEdit (GetControl ('FICHENAME')).Text;
             end;
             AGL_YFILESTD_IMPORT(FichierAIntegrer,THEdit (GetControl ('YFS_CODEPRODUIT')).Text,THEdit (GetControl ('YFS_NOM')).Text, THEdit (GetControl ('YFS_EXTFILE')).Text,
                THEdit (GetControl ('YFS_CRIT1')).Text, THEdit (GetControl ('YFS_CRIT2')).Text, THEdit (GetControl ('YFS_CRIT3')).Text,
                THEdit (GetControl ('YFS_CRIT4')).Text, THEdit (GetControl ('YFS_CRIT5')).Text,'-','-','-','-','-',
                V_PGI.LanguePrinc,
                THValComboBox (GetControl ('YFS_PREDEFINI')).Value,
                THEdit (GetControl ('YFS_LIBELLE')).Text, Numdos);
             if ExisteS  then
              DeleteFile(FichierAIntegrer);
              if OkRef then  // fiche  20862
              begin
                  DeleteFile(DocAintegrer); DocAintegrer := '';
                  OkRef := FALSE;
              end;

       end;
       FTobFts.GetEcran( Ecran ) ;
       if (ctxStandard in V_PGI.PGIContexte) then
        FTobFts.PutValue ('YFS_NODOSSIER', V_PGI.NoDossier);
       FTobFts.PutValue ('YFS_LANGUE', V_PGI.LanguePrinc);
       FTobFts.PutValue('YFS_BLOB',  RichToString( THRichEditOLE(GetControl('YFS_BLOB')) ) );
       FTobFts.InsertOrUpdateDB(TRUE);
       Ecran.ModalResult := MrOK;
       TFVierge(Ecran).BFermeClick(nil);
end;


procedure TOF_YFILESTDXL.OnNew ;
var
value : string;
begin
  Inherited ;
               SetControlText ('YFS_NOM', '');
               SetControlText ('FICHENAME', '');
               SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
               SetControlText ('YFS_CRIT1', SType);
               SetControlText ('YFS_LANGUE', V_PGI.LanguePrinc);
               if EstSpecif('51502') then
                  Value := 'CEG'
               else
               if (ctxStandard in V_PGI.PGIContexte) then
                  Value := 'STD'
               else
                  Value := 'DOS';

               SetControlText ('YFS_PREDEFINI', Value);
               THValComboBox(GetControl('YFS_PREDEFINI')).Value := Value;

               SetControlText ('YFS_EXTFILE', 'XLS');

               SetControlText ('YFS_NODOSSIER', V_PGI.NoDossier);
               SetControlText ('YFS_CRIT2', '');
               SetControlText ('YFS_LIBELLE', '');
               SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
               SetControlText ('YFS_CRIT1', SType);
               SetControlText ('YFS_CRIT3', '');

end ;

procedure TOF_YFILESTDXL.BIBLIOnClick(Sender: TObject);
var
CodeRetour : integer;
Q          : TQuery;
Value      : string;
FileName   : string;
begin
   LookupList(TControl(THEdit(Ecran.FindComponent('FICHENAME'))),'','','YFS_NOM','YFS_NOM','','',True,-1,'SELECT YFS_NOM,YFS_LIBELLE FROM YFILESTD WHERE  YFS_CRIT1="'+SType+'" AND YFS_NODOSSIER="000000" AND (YFS_PREDEFINI="CEG" OR YFS_PREDEFINI="STD")');
   Q := Opensql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+THEdit (GetControl ('FICHENAME')).Text+'" AND YFS_CRIT1="'+SType+'" AND YFS_NODOSSIER="000000" AND (YFS_PREDEFINI="CEG" OR YFS_PREDEFINI="STD")', TRUE);
   if not Q.EOF then
   begin
        CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring
        ,Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, '000000');
        if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier');
        begin
          // fiche  20862 FileName := AGL_YFILESTD_GET_PATH('COMPTA', Q.FindField ('YFS_NOM').asstring, THEdit (GetControl ('YFS_CRIT1')).Text, THEdit (GetControl ('YFS_CRIT2')).Text, THEdit (GetControl ('YFS_CRIT3')).Text, THEdit (GetControl ('YFS_CRIT4')).Text, THEdit (GetControl ('YFS_CRIT5')).Text, V_PGI.LanguePrinc, THValComboBox (GetControl ('YFS_PREDEFINI')).Value, V_PGI.NoDossier);
          FileName := TcbpPath.GetCegidUserTempPath  + ExtractFileName(FichierAIntegrer);
          {Copie du fichier extrait dans le dossier 000000 au niveau du dossier choisi}
          Copyfile (PChar( FichierAIntegrer), Pchar(FileName),TRUE);
          DocAintegrer := FileName;
          OkRef := TRUE;
          DeleteFile(FichierAIntegrer);
        end;

        SetControlText ('YFS_NOM', Q.FindField ('YFS_NOM').asstring);
        SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
        SetControlText ('YFS_CRIT1', SType);
        SetControlText ('YFS_LANGUE', V_PGI.LanguePrinc);
        if EstSpecif('51502') then
                  Value := 'CEG'
        else
        if (ctxStandard in V_PGI.PGIContexte) then
                  Value := 'STD'
        else
                  Value := 'DOS';

        SetControlText ('YFS_PREDEFINI', Value);
        SetControlText ('YFS_EXTFILE', 'XLS');
        SetControlText ('YFS_NODOSSIER', V_PGI.NoDossier);
        SetControlText ('YFS_CRIT2', Q.FindField ('YFS_CRIT2').asstring);
        SetControlText ('YFS_CRIT3', Q.FindField ('YFS_CRIT3').asstring);
        SetControlText ('YFS_CRIT4', Q.FindField ('YFS_CRIT4').asstring);
        SetControlText ('YFS_LIBELLE', Q.FindField ('YFS_LIBELLE').asstring);
        SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
        SetControlText ('YFS_CRIT1', SType);
   end;
   Ferme (Q);
end;

procedure TOF_YFILESTDXL.BREFOnClick(Sender: TObject);
var
FileName       : string;
c1,c2,c3,c4,c5 : string;
CodeRetour     : integer;
Q              : TQuery;
Value          : string;
begin
  SetControlText ('FICHENAME', 'DOC DE BASE.xls');
  {dans ce cas on va chercher dans le fichier std cegid dans db000000 pour PCL}
  Q := Opensql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+THEdit (GetControl ('FICHENAME')).Text+'" AND YFS_CRIT1="EXL" AND YFS_NODOSSIER="000000" AND (YFS_PREDEFINI="CEG" OR YFS_PREDEFINI="STD")', TRUE);
  if not Q.EOF then
  begin
      c1 := Q.FindField ('YFS_CRIT1').asstring; c2 := Q.FindField ('YFS_CRIT2').asstring;
      c3 := Q.FindField ('YFS_CRIT3').asstring; c4 := Q.FindField ('YFS_CRIT4').asstring;
      c5 := Q.FindField ('YFS_CRIT5').asstring;
      CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, c1, c2, c3
        , c4, c5, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, '000000');
        if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier')
        else
        begin
          //fiche  20862  pb si repertoire non créé FileName := AGL_YFILESTD_GET_PATH('COMPTA', Q.FindField ('YFS_NOM').asstring, THEdit (GetControl ('YFS_CRIT1')).Text, THEdit (GetControl ('YFS_CRIT2')).Text, THEdit (GetControl ('YFS_CRIT3')).Text, THEdit (GetControl ('YFS_CRIT4')).Text, THEdit (GetControl ('YFS_CRIT5')).Text, V_PGI.LanguePrinc, THValComboBox (GetControl ('YFS_PREDEFINI')).Value, V_PGI.NoDossier);
          FileName := TcbpPath.GetCegidUserTempPath  + ExtractFileName(FichierAIntegrer);
          {Copie du fichier extrait dans le dossier 000000 au niveau du dossier choisi}
          Copyfile (PChar( FichierAIntegrer), Pchar(FileName),TRUE);
          DocAintegrer := FileName;
          OkRef := TRUE;
          DeleteFile(FichierAIntegrer);
        end;
  end
  else
  begin
    PGIInfo ('Le fichier de référence inexistant dans la base');   Ferme (Q); exit;
  end;
  Ferme (Q);

   if CodeRetour = -1 then
   begin
        SetControlText ('YFS_NOM', ExtractFileName(FileName));
        SetControlText ('FICHENAME', ExtractFileName(FileName));
        SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
        SetControlText ('YFS_CRIT1', SType);
        SetControlText ('YFS_LANGUE', V_PGI.LanguePrinc);
        if EstSpecif('51502') then
                  Value := 'CEG'
        else
        if (ctxStandard in V_PGI.PGIContexte) then
                  Value := 'STD'
        else
                  Value := 'DOS';

        SetControlText ('YFS_PREDEFINI', Value);
        SetControlText ('YFS_EXTFILE', 'XLS');
        SetControlText ('YFS_NODOSSIER', V_PGI.NoDossier);
        SetControlText ('YFS_CRIT1', c1);
        SetControlText ('YFS_CRIT2', c2);
        SetControlText ('YFS_CRIT3', c3);
        SetControlText ('YFS_CRIT4', '');
        SetControlText ('YFS_CRIT5', '');
        SetControlText ('YFS_LIBELLE', 'Document de base CEGID');
        SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
        SetControlText ('YFS_CRIT1', SType);
   end;
end;

procedure TOF_YFILESTDXL.DupliqueOnClick(Sender: TObject);
var
CodeRetour : integer;
Q          : TQuery;
Value      : string;
FName,Ext  : string;
ind        : integer;
Filename   : string;
lOBM       : TOB;
FFile      : string;
i          : integer;
begin
   FFile := ExtractFileName (THEdit (GetControl ('FICHENAME')).Text);
   Ext := FFile;
   Fname   := ReadTokenPipe (Ext, '.');
   Ext := '.'+Ext;
   Ind := 1;
   i :=  Pos ('_', Fname);
   if i > 0 then
   Fname := copy(Fname, 0, i-1);

   While (ExisteSql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+Fname+'_'+ IntTostr(Ind)+Ext+'" AND YFS_CRIT1="'+SType+
   '" AND YFS_PREDEFINI="'+THValComboBox(GetControl('YFS_PREDEFINI')).Value+'" AND YFS_NODOSSIER="'+ V_PGI.NoDossier+'"')) do inc (Ind);
   Fname   := Fname+ '_'+ IntTostr(Ind)+Ext;

   Q := Opensql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+THEdit (GetControl ('FICHENAME')).Text+'" AND YFS_CRIT1="'+SType+
   '" AND YFS_PREDEFINI="'+THValComboBox(GetControl('YFS_PREDEFINI')).Value+'" AND YFS_NODOSSIER="'+ V_PGI.NoDossier+'"', TRUE);
   if not Q.EOF then
   begin
        CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring
        ,Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, V_PGI.NoDossier);
        if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier');

         if CodeRetour = -1 then
         begin

              Filename := AGL_YFILESTD_GET_PATH(Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring,
              Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring,
              Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, V_PGI.NoDossier);

              Fname := ExtractFileDir(Filename)+'\'+ Fname;
              Copyfile (PChar( Filename), Pchar(Fname),TRUE);


              SetControlText ('YFS_NOM', ExtractFileName(Fname));
              SetControlEnabled ('FICHENAME', TRUE);
              SetControlText ('FICHENAME', ExtractFileName(Fname));
              SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
              SetControlText ('YFS_CRIT1', SType);
              SetControlText ('YFS_LANGUE', V_PGI.LanguePrinc);
              if EstSpecif('51502') then
                        Value := 'CEG'
              else
              if (ctxStandard in V_PGI.PGIContexte) then
                        Value := 'STD'
              else
                        Value := 'DOS';

              SetControlText ('YFS_PREDEFINI', Value);
              SetControlText ('YFS_EXTFILE', 'XLS');
              SetControlText ('YFS_NODOSSIER', V_PGI.NoDossier);
              SetControlText ('YFS_CRIT1', Q.FindField ('YFS_CRIT1').asstring);
              SetControlText ('YFS_CRIT2', Q.FindField ('YFS_CRIT2').asstring);
              SetControlText ('YFS_CRIT3', Q.FindField ('YFS_CRIT3').asstring);
              SetControlText ('YFS_CRIT4', Q.FindField ('YFS_CRIT4').asstring);
              SetControlText ('YFS_CRIT4', Q.FindField ('YFS_CRIT5').asstring);
              SetControlText ('YFS_LIBELLE', Q.FindField ('YFS_LIBELLE').asstring);
              SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
              SetControlText ('YFS_CRIT1', SType);
              AGL_YFILESTD_IMPORT(FichierAIntegrer,THEdit (GetControl ('YFS_CODEPRODUIT')).Text,ExtractFileName(Fname), THEdit (GetControl ('YFS_EXTFILE')).Text,
                THEdit (GetControl ('YFS_CRIT1')).Text, THEdit (GetControl ('YFS_CRIT2')).Text, THEdit (GetControl ('YFS_CRIT3')).Text,
                THEdit (GetControl ('YFS_CRIT4')).Text, THEdit (GetControl ('YFS_CRIT5')).Text,'-','-','-','-','-',
                V_PGI.LanguePrinc,
                THValComboBox (GetControl ('YFS_PREDEFINI')).Value,
                THEdit (GetControl ('YFS_LIBELLE')).Text, V_PGI.NoDossier);
              lOBM := TOB.Create('YFILESTD', nil, -1 ) ;
              lOBM.Dupliquer( FTobFts, True, True ) ;
              lOBM.InsertOrUpdateDB(TRUE);
              lOBM.free;
         end;
   end;
   Ferme (Q);
end;

procedure TOF_YFILESTDXL.OnClose ;
begin
              if OkRef and (DocAintegrer <> '') then
              begin
                  DeleteFile(DocAintegrer); DocAintegrer := '';
                  OkRef := FALSE;
              end;
end;

{$IFDEF SCANGED}

Procedure MyImportGed (FileGuid : string; Lib : string);
var
 ParGed      : TParamGedDoc;
 NomFichier  : string;
 Annee,mm,jj : word ;
begin
  if not (ctxPCL in V_PGI.PGIContexte) then exit;
    // ---- Insertion classique dans la GED avec boite de dialogue ----
      // Propose le rangement dans la GED
      ParGed.SDocGUID  := '';
      ParGed.NoDossier := V_PGI.NoDossier;
      ParGed.CodeGed   := '';
      DecodeDate(GetEnCours.Fin , Annee, mm, jj);
      ParGed.Annee     := IntToStr(Annee);
      ParGed.Mois      := IntToStr(mm);
      ParGed.DefName   := Lib;
      // FileId est le n° de fichier obtenu par la GED suite à l'insertion
      ParGed.SFileGUID := FileGuid;

      Application.BringToFront;

      if (JaileDroitConceptBureau (187315)) then
      begin
        if ShowNewDocument(ParGed, FALSE, FALSE, FALSE, 'A')='##NONE##' then
          // Fichier refusé, suppression dans la GED
          V_GedFiles.Erase(FileGuid);
      end
      else
      begin
        if PgiAsk ('Vous n''avez pas le droit d''insérer un document dans la GED'#10' Voulez-vous enregistrer ce document sur disque?') = mrYes then
        with TSaveDialog.Create (nil) do
        begin
          Options    := [ofOverwritePrompt];
          FileName   := GetFileNameGed (FileGUID);
          Nomfichier := ExtractFileExt (FileName);
          Filter     := 'Fichiers *' + NomFichier +'|*' + NomFichier + '|Tous les fichiers (*.*)|*.*';
          if Execute = TRUE then
          if V_GedFiles.Extract (FileName, FileGUID) = TRUE then
            PgiInfo ('Le document a été enregistré: ' + FileName)
          else
            PgiInfo ('Impossible d''enregistrer le document sous: ' + FileName);
            Free;
        end;
        V_GedFiles.Erase(FileGUID);
      end;
end;

{$ENDIF}


Initialization
  registerclasses ( [ TOF_YFILESTDXL ] ) ;

end.
