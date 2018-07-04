unit UIMPORTBOB_TOM;

interface

Uses StdCtrls, Controls, Classes, Windows,
{$IFDEF EAGLCLIENT}
     eFiche,eFichList, Maineagl,
{$ELSE}
     db,{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}Fiche, FichList,  Fe_Main,
{$ENDIF}
     forms, sysutils,ComCtrls,HCtrls, HEnt1, HMsgBox, UTOM, dialogs, 
     UTob, HPanel, uYFILESTD, HTB97, Lookup, Ent1, ULibWindows, Vierge;


Type
  TOM_YFILESTD = Class (TOM)
   procedure OnArgument(stArgument: String); override ;
   procedure OnUpdateRecord                ; override ;
   procedure OnLoadRecord                  ; override ;
   procedure OnChangeField ( F: TField )   ; override ;
   procedure OnNewRecord                   ; override ;
   procedure OnAfterUpdateRecord           ; override ;
   private
   SType, FichierAIntegrer                 : string;
   LeMode                                  : TActionFiche ;
   Pages                                   : TPageControl;
   SaisieCommentaire                       : Boolean;
   procedure BIBLIOnClick(Sender: TObject);
   procedure BREFOnClick(Sender: TObject);
   procedure DupliqueOnClick(Sender: TObject);


end;
implementation

procedure TOM_YFILESTD.OnArgument(stArgument: String);
var
St : string;
begin
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
        FichierAIntegrer := '';
end;

procedure TOM_YFILESTD.OnLoadRecord ;
var
Nom       : string;
lStVal    : string;
begin
        Nom := THEdit (GetControl ('YFS_NOM')).Text;
        Pages := TPageControl(GetControl('PAGES', true));

        SetControlText ('FICHENAME', Nom);
       if (SType <> '') then
       begin
             if ((DS.State in [dsInsert]) or
              (LeMode in [taCreat..taCreatOne])) then
              begin
                     TFFiche(Ecran).TypeAction := TaCreat;
                     OnNewRecord;
                     if EstSpecif('51502') then
                        lStVal := 'CEG'
                     else
                     if (ctxStandard in V_PGI.PGIContexte) then
                        lStVal := 'STD'
                     else
                        lStVal := 'DOS';

                     SetField ('YFS_PREDEFINI', lStVal);
                     SetControlEnabled ('YFS_PREDEFINI', FALSE);
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
            if SaisieCommentaire then
            begin
               Pages.Pages[0].TabVisible := FALSE;
               Pages.Pages[1].TabVisible := TRUE;
            end
            else
            begin
               Pages.Pages[1].TabVisible := FALSE;
               Pages.Pages[0].TabVisible := TRUE;
               SetControlEnabled ('YFS_CRIT1', FALSE);
               SetControlEnabled ('YFS_CRIT2', TRUE);
               SetControlEnabled ('YFS_CRIT3', TRUE);
               SetControlEnabled ('YFS_CRIT4', TRUE);
               SetControlEnabled ('YFS_CRIT5', TRUE);
            end;

            lStVal := FormatDateTime( 'dd mmm yyyy', GetField('YFS_DATEMODIF') ) ;
            SetControlText('DATEMODIF', lStVal ) ;
            lStVal := FormatDateTime( 'dd mmm yyyy', GetField('YFS_DATECREATION') ) ;
            SetControlText('DATECREATION', lStVal ) ;
      end;
end;

procedure TOM_YFILESTD.OnUpdateRecord;
begin
end;

procedure TOM_YFILESTD.OnAfterUpdateRecord;
var
Numdos  : string;
begin
        SetControlText ('YFS_NOM', UpperCase(ExtractFileName(THEdit (GetControl ('FICHENAME')).Text)));
        if Length (THEdit (GetControl ('YFS_NOM')).Text) > 35 then
        begin
             PgiInfo ('Attention , la longueur du fichier ne doit pas dépasser de 35 caractères');  LastError := -1; exit;
        end;

        if EstSpecif('51502') then
                Numdos := '000000'
        else
                Numdos := V_PGI.NoDossier;

        if ((SType <> '') and (LeMode = taCreat)) or (SType = '')  then
        begin
             if FichierAIntegrer = '' then   FichierAIntegrer := THEdit (GetControl ('FICHENAME')).Text;
             AGL_YFILESTD_IMPORT(FichierAIntegrer,THEdit (GetControl ('YFS_CODEPRODUIT')).Text,THEdit (GetControl ('YFS_NOM')).Text, THEdit (GetControl ('YFS_EXTFILE')).Text,
                THEdit (GetControl ('YFS_CRIT1')).Text, THEdit (GetControl ('YFS_CRIT2')).Text, THEdit (GetControl ('YFS_CRIT3')).Text,
                THEdit (GetControl ('YFS_CRIT4')).Text, THEdit (GetControl ('YFS_CRIT5')).Text,'-','-','-','-','-',
                V_PGI.LanguePrinc,
                THValComboBox (GetControl ('YFS_PREDEFINI')).Value,
                THEdit (GetControl ('YFS_LIBELLE')).Text, Numdos);
       end;
       Ecran.ModalResult := MrOK;
       TFVierge(Ecran).BFermeClick(nil);
end;

procedure TOM_YFILESTD.OnChangeField ( F: TField ) ;
begin
end;

procedure TOM_YFILESTD.OnNewRecord ;
var
value : string;
begin
  Inherited ;
               SetField ('YFS_NOM', '');
               SetControlText ('FICHENAME', '');
               SetField ('YFS_CODEPRODUIT', 'COMPTA');
               SetField ('YFS_CRIT1', SType);
               SetField ('YFS_LANGUE', V_PGI.LanguePrinc);
               if EstSpecif('51502') then
                  Value := 'CEG'
               else
               if (ctxStandard in V_PGI.PGIContexte) then
                  Value := 'STD'
               else
                  Value := 'DOS';

               SetField ('YFS_PREDEFINI', Value);
               THValComboBox(GetControl('YFS_PREDEFINI')).Value := Value;

               SetField ('YFS_EXTFILE', 'XLS');

               SetField ('YFS_NODOSSIER', V_PGI.NoDossier);
               SetField ('YFS_CRIT2', '');
               SetField ('YFS_LIBELLE', '');
               SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
               SetControlText ('YFS_CRIT1', SType);

end ;

procedure TOM_YFILESTD.BIBLIOnClick(Sender: TObject);
var
CodeRetour : integer;
Q          : TQuery;
Value      : string;
begin
   LookupList(TControl(THEdit(Ecran.FindComponent('FICHENAME'))),'','','YFS_NOM','YFS_NOM','','',True,-1,'SELECT YFS_NOM,YFS_LIBELLE FROM YFILESTD WHERE  YFS_CRIT1="'+SType+'" AND YFS_NODOSSIER="000000" AND (YFS_PREDEFINI="CEG" OR YFS_PREDEFINI="STD")');
   Q := Opensql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+THEdit (GetControl ('FICHENAME')).Text+'" AND YFS_CRIT1="'+SType+'" AND YFS_NODOSSIER="000000" AND (YFS_PREDEFINI="CEG" OR YFS_PREDEFINI="STD")', TRUE);
   if not Q.EOF then
   begin
        FichierAIntegrer :=  V_PGI.DosPath +Q.FindField ('YFS_NOM').asstring;
        CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring
        ,Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring);
        if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier');

        SetField ('YFS_NOM', Q.FindField ('YFS_NOM').asstring);
        SetField ('YFS_CODEPRODUIT', 'COMPTA');
        SetField ('YFS_CRIT1', SType);
        SetField ('YFS_LANGUE', V_PGI.LanguePrinc);
        if EstSpecif('51502') then
                  Value := 'CEG'
        else
        if (ctxStandard in V_PGI.PGIContexte) then
                  Value := 'STD'
        else
                  Value := 'DOS';

        SetField ('YFS_PREDEFINI', Value);
        SetField ('YFS_EXTFILE', 'XLS');
        SetField ('YFS_NODOSSIER', V_PGI.NoDossier);
        SetField ('YFS_CRIT2', Q.FindField ('YFS_CRIT2').asstring);
        SetField ('YFS_CRIT3', Q.FindField ('YFS_CRIT3').asstring);
        SetField ('YFS_CRIT4', Q.FindField ('YFS_CRIT4').asstring);
        SetField ('YFS_LIBELLE', Q.FindField ('YFS_LIBELLE').asstring);
        SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
        SetControlText ('YFS_CRIT1', SType);
   end;
   Ferme (Q);

end;

procedure TOM_YFILESTD.BREFOnClick(Sender: TObject);
var
FileName       : string;
c1,c2,c3,c4,c5 : string;
CodeRetour     : integer;
Q              : TQuery;
Value          : string;
begin
   Filename := V_PGI.StdPath + '\DOC DE BASE.xls';
   CodeRetour := -1;
   if FileExists (Filename) then
       FichierAIntegrer := FileName
   else
   begin
        Q := Opensql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+THEdit (GetControl ('FICHENAME')).Text+'" AND YFS_CRIT1="FTS" AND YFS_NODOSSIER="000000" AND (YFS_PREDEFINI="CEG" OR YFS_PREDEFINI="STD")', TRUE);
        if not Q.EOF then
        begin
            FichierAIntegrer :=  V_PGI.DosPath +Q.FindField ('YFS_NOM').asstring;
            c1 := Q.FindField ('YFS_CRIT1').asstring; c2 := Q.FindField ('YFS_CRIT2').asstring;
            c3 := Q.FindField ('YFS_CRIT3').asstring; c4 := Q.FindField ('YFS_CRIT4').asstring;
            c5 := Q.FindField ('YFS_CRIT5').asstring;
            CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, c1, c2, c3
              , c4, c5, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring);
              if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier');
        end;
   end;
   if CodeRetour = -1 then
   begin
        SetField ('YFS_NOM', ExtractFileName(FileName));
        SetControlText ('FICHENAME', ExtractFileName(FileName));
        SetField ('YFS_CODEPRODUIT', 'COMPTA');
        SetField ('YFS_CRIT1', SType);
        SetField ('YFS_LANGUE', V_PGI.LanguePrinc);
        if EstSpecif('51502') then
                  Value := 'CEG'
        else
        if (ctxStandard in V_PGI.PGIContexte) then
                  Value := 'STD'
        else
                  Value := 'DOS';

        SetField ('YFS_PREDEFINI', Value);
        SetField ('YFS_EXTFILE', 'XLS');
        SetField ('YFS_NODOSSIER', V_PGI.NoDossier);
        SetField ('YFS_CRIT1', c1);
        SetField ('YFS_CRIT2', c2);
        SetField ('YFS_CRIT3', c3);
        SetField ('YFS_CRIT4', c4);
        SetField ('YFS_CRIT4', c5);
        SetField ('YFS_LIBELLE', 'Document de base CEGID');
        SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
        SetControlText ('YFS_CRIT1', SType);
   end;
end;

procedure TOM_YFILESTD.DupliqueOnClick(Sender: TObject);
var
CodeRetour : integer;
Q          : TQuery;
Value      : string;
FName,Ext  : string;
ind        : integer;
Filename   : string;
begin

   Ext := ExtractFileName (THEdit (GetControl ('FICHENAME')).Text);
   Fname   := ReadTokenPipe (Ext, '.');
   Ext := '.'+Ext;
   Ind := 1;
   While (ExisteSql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+Fname+ IntTostr(Ind)+Ext+'" AND YFS_CRIT1="'+SType+
   '" AND YFS_PREDEFINI="DOS" AND YFS_NODOSSIER="'+ V_PGI.NoDossier+'"')) do inc (Ind);
   Fname   := Fname+ IntTostr(Ind)+Ext;

   Q := Opensql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+THEdit (GetControl ('FICHENAME')).Text+'" AND YFS_CRIT1="'+SType+
   '" AND YFS_PREDEFINI="DOS" AND YFS_NODOSSIER="'+ V_PGI.NoDossier+'"', TRUE);
   if not Q.EOF then
   begin
        FichierAIntegrer :=  V_PGI.DosPath +Q.FindField ('YFS_NOM').asstring;
        CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring
        ,Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring);
        if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier');

         if CodeRetour = -1 then
         begin
              Filename := GetWindowsTempPath +'PGI\STD\'+Q.FindField ('YFS_CODEPRODUIT').asstring+'\';
              if Q.FindField ('YFS_CRIT1').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT1').asstring+'\';
              if Q.FindField ('YFS_CRIT2').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT2').asstring+'\';
              if Q.FindField ('YFS_CRIT3').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT3').asstring+'\';
              if Q.FindField ('YFS_CRIT4').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT4').asstring+'\';
              if Q.FindField ('YFS_CRIT5').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT5').asstring+'\';
              Filename := Filename + Q.FindField ('YFS_LANGUE').asstring+'\'+Q.FindField ('YFS_PREDEFINI').asstring+'\'+
              Q.FindField ('YFS_NOM').asstring;

              Fname := ExtractFileDir(Filename)+'\'+ Fname;
              Copyfile (PChar( Filename), Pchar(Fname),TRUE);


              SetField ('YFS_NOM', ExtractFileName(Fname));
              SetControlEnabled ('FICHENAME', TRUE);
              SetControlText ('FICHENAME', ExtractFileName(Fname));
              SetField ('YFS_CODEPRODUIT', 'COMPTA');
              SetField ('YFS_CRIT1', SType);
              SetField ('YFS_LANGUE', V_PGI.LanguePrinc);
              if EstSpecif('51502') then
                        Value := 'CEG'
              else
              if (ctxStandard in V_PGI.PGIContexte) then
                        Value := 'STD'
              else
                        Value := 'DOS';

              SetField ('YFS_PREDEFINI', Value);
              SetField ('YFS_EXTFILE', 'XLS');
              SetField ('YFS_NODOSSIER', V_PGI.NoDossier);
              SetField ('YFS_CRIT1', Q.FindField ('YFS_CRIT1').asstring);
              SetField ('YFS_CRIT2', Q.FindField ('YFS_CRIT2').asstring);
              SetField ('YFS_CRIT3', Q.FindField ('YFS_CRIT3').asstring);
              SetField ('YFS_CRIT4', Q.FindField ('YFS_CRIT4').asstring);
              SetField ('YFS_CRIT4', Q.FindField ('YFS_CRIT5').asstring);
              SetField ('YFS_LIBELLE', Q.FindField ('YFS_LIBELLE').asstring);
              SetControlText ('YFS_CODEPRODUIT', 'COMPTA');
              SetControlText ('YFS_CRIT1', SType);
              AGL_YFILESTD_IMPORT(FichierAIntegrer,THEdit (GetControl ('YFS_CODEPRODUIT')).Text,ExtractFileName(Fname), THEdit (GetControl ('YFS_EXTFILE')).Text,
                THEdit (GetControl ('YFS_CRIT1')).Text, THEdit (GetControl ('YFS_CRIT2')).Text, THEdit (GetControl ('YFS_CRIT3')).Text,
                THEdit (GetControl ('YFS_CRIT4')).Text, THEdit (GetControl ('YFS_CRIT5')).Text,'-','-','-','-','-',
                V_PGI.LanguePrinc,
                THValComboBox (GetControl ('YFS_PREDEFINI')).Value,
                THEdit (GetControl ('YFS_LIBELLE')).Text, V_PGI.NoDossier);

         end;
   end;
   Ferme (Q);

end;

Initialization
  registerclasses ( [ TOM_YFILESTD ] ) ;

end.
