unit UIMPORTBOB_TOF;

interface

uses StdCtrls, Classes, Windows, UTOF, HStatus, M3FP, Sysutils, Shellapi,
{$IFDEF EAGLCLIENT}
     MaineAgl,
     eMul,
{$ELSE}
     Mul, HDB, Db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_main,
{$ENDIF}
  AglInit,
  HEnt1, HTB97, Dialogs, Forms,Controls, HQry,
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  HMsgBox, HCTRLS, Lookup, uYFILESTD, HSysMenu, Ent1, ULibWindows, Vierge, UTOB, UYFILESTD_TOF;


type
  TOF_IMPORTBOB = class(TOF)
    procedure OnNew; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose ; override ;

    private
    SType     : string;
    Numdos    : string;
    Q         : TQuery;
    FEcran    : TFMul ;
{$IFDEF EAGLCLIENT}
    FListe    : THGrid;
    wBookMark : integer;
{$ELSE}
    FListe    : THDBGrid;
    wBookMark : tBookMark;
{$ENDIF}
    procedure BCOMMENTAIREOnClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure SupprOnClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
    procedure SupprimeEnreg(L : THGrid; Q : THQuery);
{$ELSE}
    procedure SupprimeEnreg(L : THDBGrid; Q : THQuery);
{$ENDIF}
    procedure BDblOnClick(Sender: TObject);
    procedure BREINTEGEROnClick(Sender: TObject);
    function UpdateEnreg(Q : TQuery) : Boolean;
    procedure BOnClick(Sender: TObject);
{$IFDEF SCANGED}
    procedure GedOnClick(Sender: TObject);
{$ENDIF}
    procedure SaveRow;
    procedure RestorRow;
  //  Function RendCommandeCRIT(ts : string) : string;
    procedure DupliqueOnClick(Sender: TObject);
    procedure PredefiniChange(Sender: TObject);

  end;

procedure CPLanceFiche_CPIMPFICHEXCEL (TypS : string);

implementation


procedure CPLanceFiche_CPIMPFICHEXCEL (TypS : string);
begin
    AglLanceFiche('CP','CPIMPFICHEXCEL','', '',  TypS)  ;
end;


procedure TOF_IMPORTBOB.OnNew;
begin
  inherited;
end;

procedure TOF_IMPORTBOB.OnUpdate;
begin
  inherited;
end;

procedure TOF_IMPORTBOB.OnLoad;
begin
  inherited;
   if not (ctxPcl in V_PGI.PGIContexte) then
     THValComboBox (GetControl ('YFS_PREDEFINI')).Plus :=  ' AND CO_CODE<>"STD"';
   if SType = 'EXL' then
          Ecran.Caption := 'Autres feuilles Excel'
   else
   if SType = 'FTS' then
          Ecran.Caption := 'Feuilles de travail spécialisées';

   UpdateCaption( Ecran ) ;
end;
(*
Function TOF_IMPORTBOB.RendCommandeCRIT(ts : string) : string;
var
St       : string;
x        : string;
begin
St := '';
x := ReadTokenSt (ts);
if x <> '' then  St := 'YFS_CRIT3="' +x+ '"';
while x <> '' do
  begin
       x := ReadTokenSt (ts);
       if x <> '' then
       St := St+ ' OR '+ 'YFS_CRIT3="' +x+ '"';
  end;
  Result := St;
end;
*)

procedure TOF_IMPORTBOB.OnArgument(S: string);
var
WWW, St       : string;
begin
  FEcran := TFMul(Ecran) ;
  // fiche 20402
  Fecran.FiltreDisabled := true ;

  inherited;


{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FListe'));
{$ELSE}
    FListe := THDBGrid (GetControl('FListe'));
    CentreDBGrid(FListe);
{$ENDIF}

  TToolbarButton97(GetControl('BINSERT')).OnClick        := BInsertClick;
  TToolBarButton97(GetControl('BDelete')).OnClick        := SupprOnClick ;
  TToolBarButton97(GetControl('DUPLIQUER')).OnClick      := DupliqueOnClick;
{$IFDEF SCANGED}
  if (ctxPCL in V_PGI.PGIContexte) then
    TToolBarButton97(GetControl('BGED')).OnClick         := GedOnClick
  else
    TToolBarButton97(GetControl('BGED')).visible         := FALSE;
{$ELSE}
    TToolBarButton97(GetControl('BGED')).visible         := FALSE;
{$ENDIF}
{$IFDEF EAGLCLIENT}
  FEcran.HMTrad.ResizeGridColumns(FListe);
{$ELSE}
  THSystemMenu(GetControl('HMTrad')).ResizeDBGridColumns(FListe);
{$ENDIF}

  // GCO - 02/10/2007 - FQ 21558
  THEdit(GetControl ('YFS_CRIT3')).Plus := 'AND CCY_EXERCICE="'+ GetEncours.Code + '"';

  if S <> '' then
  begin
       St := S;
       SType := ReadTokenSt(St);
       if St <> '' then
       begin
          WWW := ReadTokenSt(St);
          SetControlText ('YFS_CRIT2', WWW);
          SetControlText ('YFS_CRIT2_', WWW);
       end;
       if St <> '' then
       begin
          //StCrit3 := RendCommandeCRIT (St);
          SetControlText ('YFS_CRIT3', St);
       end;
       FListe.OnDblClick                                      := BDblOnClick;
       TToolBarButton97(GetControl('BOuvrir')).OnClick        := BOnClick;
       TToolBarButton97(GetControl('BCOMMENTAIRE')).OnClick   := BCOMMENTAIREOnClick;
       TToolBarButton97(GetControl('BREINTEGER')).OnClick     := BREINTEGEROnClick;
       if EstSpecif('51502') then
       begin
          WWW := 'CEG';  Numdos := '000000';
       end
       else
       begin
            if (ctxStandard in V_PGI.PGIContexte) then
            begin
                WWW := 'STD'; Numdos := '000000';
            end
            else
            begin
                WWW := 'DOS'; Numdos := V_PGI.NoDossier;
            end;
       end;
       THValComboBox(GetControl('YFS_PREDEFINI')).Value :=  WWW;
       if WWW = 'STD' then   // fiche 21557
       begin
          THValComboBox(GetControl('YFS_PREDEFINI')).Plus := 'AND CO_CODE<>"DOS"';
          if SType = 'EXL' then THValComboBox(GetControl('YFS_PREDEFINI')).value := '';
       end;
       if ((WWW = 'CEG') and EstSpecif('51502')) or (WWW = 'STD') then
       begin
          SetControlEnabled ('YFS_PREDEFINI', TRUE);
          if ExisteSQL ('SELECT YFS_BCRIT2 FROM YFILESTD Where YFS_BCRIT2="X" AND YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="'+SType+'" AND (YFS_PREDEFINI="'+WWW+'")') then
             ExecuteSQL ('UPDATE YFILESTD SET YFS_BCRIT2="-" Where YFS_CODEPRODUIT="COMPTA" AND YFS_PREDEFINI="'+ WWW + '" AND YFS_CRIT1="'+SType+'"');
       end
       else
       begin
        if ExisteSQL ('SELECT YFS_BCRIT2 FROM YFILESTD Where YFS_BCRIT2="-" AND YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="'+SType+'" AND (YFS_PREDEFINI="STD" or YFS_PREDEFINI="CEG")') then
          ExecuteSQL ('UPDATE YFILESTD SET YFS_BCRIT2="X" Where YFS_CODEPRODUIT="COMPTA" AND YFS_BCRIT2<>"X" AND YFS_CRIT1="'+SType+'" AND (YFS_PREDEFINI="STD" or YFS_PREDEFINI="CEG")');
       end;

       THEdit(GetControl('XX_WHERE')).Text := 'YFS_CRIT1="' + SType + '" AND' +
       ' YFS_NODOSSIER="'+ Numdos +'"';

(*      if  getControlText ('YFS_CRIT2') <> '' then
       THEdit(GetControl('XX_WHERE')).Text := THEdit(GetControl('XX_WHERE')).Text +
        ' AND YFS_CRIT2="'+getControlText ('YFS_CRIT2')+'"'
       else
       if  (getControlText ('YFS_CRIT3') <> '') and (getControlText ('YFS_CRIT3') <> '<<Tous>>') then
       THEdit(GetControl('XX_WHERE')).Text := THEdit(GetControl('XX_WHERE')).Text +
        ' AND ('+ StCrit3+')';
*)
   end;
   THValComboBox (GetControl ('YFS_PREDEFINI')).OnChange := PredefiniChange;
{$IFDEF EAGLCLIENT}
  FEcran.HMTrad.ResizeGridColumns(FListe);
{$ELSE}
  FEcran.HMTrad.ResizeDBGridColumns(FListe);
{$ENDIF}
// Affichage des zones dans le multicritères
  SetControlVisible('TE_GENERAL',SType='FTS');
  SetControlVisible('TE_GENERAL_',SType='FTS');
  SetControlVisible('YFS_CRIT2',SType='FTS');
  SetControlVisible('YFS_CRIT2_',SType='FTS');
  SetControlVisible('TYFS_CRIT3',SType='FTS');
  SetControlVisible('YFS_CRIT3',SType='FTS');
  if SType='EXL' then TFMul(Ecran).SetDBListe('CPFICHIEREXCEL');
end;

procedure TOF_IMPORTBOB.BInsertClick(Sender: TObject);
var
WWW       : string;
begin
inherited;
       if EstSpecif('51502') then
          WWW := 'CEG'
       else
       if (ctxStandard in V_PGI.PGIContexte) then
          WWW := 'STD'
       else
          WWW := 'DOS';

     if (TheTob = nil) then TheTob := TOB.Create('YFILESTD', nil, -1);
 (*fiche 20686   TheTob.SelectDB('', nil, False);
     TheTob.LoadDB; // pour les memos
 *)
     SaveRow;
     if SType <> '' then
        CPLanceFiche_YFILESTDXL ('', SType+';ACTION=CREATION')
     else
        AglLanceFiche('CP','CPIMPORTFICH','', '', SType)  ;
     //ImportFichierBob (Cle)  ;
     PredefiniChange (nil);
     FEcran.BChercheClick(nil) ;
     RestorRow;
end;

{$IFDEF EAGLCLIENT}
procedure TOF_IMPORTBOB.SupprimeEnreg(L : THGrid; Q : THQuery);
{$ELSE}
procedure TOF_IMPORTBOB.SupprimeEnreg(L : THDBGrid; Q : THQuery);
{$ENDIF}
var
    i : integer;
    Stt : string;
begin
     if Q.FindField ('YFS_BCRIT2').asstring = 'X' then begin PGIInfo ('Suppression interdite'); exit; end;


     if EstSpecif('51502') or (ctxStandard in V_PGI.PGIContexte) then
                Numdos := '000000'
     else
                Numdos := V_PGI.NoDossier;

    // si rien de selectionné !
    if ((L.NbSelected = 0) and (not L.AllSelected)) then
    begin
        MessageAlerte('Aucun élément sélectionné.');
        exit;
    end;

    // message d'avertisement !
    if (PGIAsk(TraduireMemoire('Vous allez supprimer définitivement les informations. Confirmez vous l''opération ?')) <> mrYes) then exit;

    // destruction
    if (L.AllSelected) then // si tout ??
    begin
        Q.First;

        while (Not Q.EOF) do
        begin
            MoveCur(False);
            if (Q.FindField ('YFS_BCRIT1').asstring = 'X') then
                                         PgiInfo ('Attention le fichier '+ Q.FindField ('YFS_NOM').asstring + ' est extrait. Suppression interdite')
            else
            begin

                  Stt := ' WHERE YFS_CODEPRODUIT="'+Q.FindField ('YFS_CODEPRODUIT').asstring +'" AND YFS_CRIT1="'
                  + Q.FindField ('YFS_CRIT1').asstring +'" AND YFS_CRIT2="' + Q.FindField ('YFS_CRIT2').asstring +'"'
                  +' AND YFS_NOM="' + Q.FindField ('YFS_NOM').asstring +'"'
                  +' AND YFS_LANGUE="' + Q.FindField ('YFS_LANGUE').asstring +'"'
                  +' AND YFS_PREDEFINI="' + Q.FindField ('YFS_PREDEFINI').asstring +'"'
                  +' AND YFS_EXTFILE="' + Q.FindField ('YFS_EXTFILE').asstring +'"'
                  +' AND YFS_NODOSSIER="'+Numdos+'"';
                  ExecuteSQl ('DELETE FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD ' + Stt +')');
                  ExecuteSQl ('DELETE FROM YFILESTD ' + Stt);
                  ExecuteSQl ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD ' + Stt +')');
            end;
            Q.Next;
        end;

        L.AllSelected:=False;
    end
    else // si pas tout ....
    begin
        InitMove(L.NbSelected,'');

        for i := 0 to L.NbSelected-1 do
        begin
            MoveCur(False);
            L.GotoLeBookMark(i);
            {$IFDEF EAGLCLIENT}
            FEcran.Q.TQ.Seek(L.Row-1);
            {$ENDIF}
            if (Q.FindField ('YFS_BCRIT1').asstring = 'X') then
                                         PgiInfo ('Attention le fichier '+ Q.FindField ('YFS_NOM').asstring + ' est extrait. Suppression interdite')
            else
            begin
                 Stt := ' WHERE YFS_CODEPRODUIT="'+Q.FindField ('YFS_CODEPRODUIT').asstring +'" AND YFS_CRIT1="' +
                  Q.FindField ('YFS_CRIT1').asstring +'" AND YFS_CRIT2="' + Q.FindField ('YFS_CRIT2').asstring +'"'
                  +' AND YFS_NOM="' + Q.FindField ('YFS_NOM').asstring +'"'
                  +' AND YFS_LANGUE="' + Q.FindField ('YFS_LANGUE').asstring +'"'
                  +' AND YFS_PREDEFINI="' + Q.FindField ('YFS_PREDEFINI').asstring +'"'
                  +' AND YFS_EXTFILE="' + Q.FindField ('YFS_EXTFILE').asstring +'"'
                  +' AND YFS_NODOSSIER="'+Numdos+'"';
                  ExecuteSQl ('DELETE FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD ' + Stt +')');
                  ExecuteSQl ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD ' + Stt +')');
                  ExecuteSQl ('DELETE FROM YFILESTD ' + Stt);
            end;
        end;

        L.ClearSelected;
    end;
    FiniMove;
    FEcran.BChercheClick(nil) ;

end;

procedure TOF_IMPORTBOB.SupprOnClick(Sender: TObject);
var
    F : TFMul;
    Query : THQuery;
begin
    F := TFMul(Longint( Ecran ));
    if (F = Nil) then exit;

    if (FListe = Nil) then exit;

{$IFDEF EAGLCLIENT}
    FEcran.Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
    Query := F.Q;
    if (Query = Nil) then exit;
     if Query.EOF then
     begin
        PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
     end;

    SupprimeEnreg(Fliste,Query);
end;

procedure TOF_IMPORTBOB.BDblOnClick(Sender: TObject);
var
CodeRetour : integer;
Filename   : string;
rr         : integer;
begin
        SaveRow;
        if Q.EOF then
        begin
            PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
        end;

        if Q.FindField ('YFS_BCRIT2').asstring = 'X' then begin PGIInfo ('Accès interdit'); exit; end;

        Filename := AGL_YFILESTD_GET_PATH(Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring,
        Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring,
        Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, Numdos);

        if not FileExists (Filename) then
        begin
              CodeRetour :=  AGL_YFILESTD_EXTRACT (Filename, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring
              ,Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, Numdos);
              if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier')
        end
        else
        begin
                    rr := FileOpen(FileName, fmShareExclusive);
                    if (rr = -1) and (Q.FindField ('YFS_BCRIT1').asstring = 'X') then
                    begin
                                 PgiInfo ('Attention le fichier '+ Q.FindField ('YFS_NOM').asstring + ' est déjà ouvert');
                                 FileClose (rr);
                                 exit;
                    end;
                    FileClose (rr);
                    CodeRetour := -1;
        end;
        if CodeRetour = -1 then
        begin
            ExecuteSQL ('UPDATE YFILESTD SET YFS_BCRIT1="X" AND WHERE YFS_NOM="'+Q.FindField ('YFS_NOM').asstring+'" AND YFS_CRIT1="'+Q.FindField ('YFS_CRIT1').asstring+'" AND '+
            'YFS_CRIT2="'+Q.FindField ('YFS_CRIT2').asstring+'" AND ' +
            'YFS_CRIT3="'+Q.FindField ('YFS_CRIT3').asstring+'" AND ' +
            'YFS_CRIT4="'+Q.FindField ('YFS_CRIT4').asstring+'" AND ' +
            'YFS_CRIT5="'+Q.FindField ('YFS_CRIT5').asstring+'" AND ' +
            'YFS_PREDEFINI="'+Q.FindField ('YFS_PREDEFINI').asstring+'" AND ' +
            'YFS_NODOSSIER="'+ Numdos +'" AND '+
            'YFS_LANGUE="'+Q.FindField ('YFS_LANGUE').asstring+'"');
             if  Q.FindField ('YFS_EXTFILE').asstring = 'PDF' then
                     ShellExecute(0, PCHAR('open'), 'AcroRd32.exe' , PCHAR('"'+Filename+'"' ),PCHAR(Filename ),SW_RESTORE)
             else
             if  (Q.FindField ('YFS_EXTFILE').asstring = 'XLS') or (Q.FindField ('YFS_EXTFILE').asstring = 'CSV') then
                    ShellExecute(0, PCHAR('open'), 'Excel.exe' , PCHAR('"'+Filename+'"' ),PCHAR(Filename ),SW_RESTORE)
             else
             if  Q.FindField ('YFS_EXTFILE').asstring = 'TXT' then
                     ShellExecute(0, PCHAR('open'), 'WordPad.exe' , PCHAR('"'+Filename+'"' ),PCHAR(Filename ),SW_RESTORE)
        end;
        FEcran.BChercheClick(nil);
        RestorRow;
end;


procedure TOF_IMPORTBOB.BOnClick(Sender: TObject);
begin
     SaveRow;
     if Q.EOF then
     begin
        PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
     end;

     if (not Q.EOF) and (Q.FindField ('YFS_BCRIT2').asstring = 'X') then PGIInfo ('Accès uniquement en consultation');
     if (TheTob = nil) then TheTob := TOB.Create('YFILESTD', nil, -1);
     TheTob.SelectDB('', Q, False);
     if Q.EOF then
     begin
        PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
     end;
     TheTob.LoadDB; // pour les memos
     CPLanceFiche_YFILESTDXL (GetField ('YFS_NOM')+';'+GetField ('YFS_PREDEFINI')+';'+GetField ('YFS_CRIT1'), GetField ('YFS_CRIT1'));

     FEcran.BChercheClick(nil);
     RestorRow;
end;

procedure TOF_IMPORTBOB.BCOMMENTAIREOnClick(Sender: TObject);
begin
     SaveRow;
     if (TheTob = nil) then TheTob := TOB.Create('YFILESTD', nil, -1);
     TheTob.SelectDB('', Q, False);
     if Q.EOF then
     begin
        PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
     end;
     TheTob.LoadDB; // pour les memos
     if (SType = 'FTS') or (SType = 'EXL')then
     CPLanceFiche_YFILESTDXL ( GetField ('YFS_NOM')+';'+GetField ('YFS_PREDEFINI')+';'+GetField('YFS_CRIT1'), SType+';ACTION=COMMENTAIRE');

     FEcran.BChercheClick(nil);
     RestorRow;
end;


procedure TOF_IMPORTBOB.BREINTEGEROnClick(Sender: TObject);
var
    F : TFMul;
    i : integer;
begin
    F := TFMul(Longint( Ecran ));
    if (F = Nil) then exit;
    if (FListe = Nil) then exit;
    if (Fliste.NbSelected = 0)  and (not Fliste.AllSelected)then
    begin
            MessageAlerte('sélectionnez un élément.');
            exit;
    end;
    if PGIAsk ('Confirmez-vous la sauvegarde des documents dans la base ?') <> mryes then
    begin Fliste.ClearSelected; exit;   end;

    SaveRow;
    if (Q = Nil) then exit;
    if (Fliste.AllSelected) then
    begin
        Q.First;

        while (Not Q.EOF) do
        begin
            MoveCur(False);
            UpdateEnreg(Q);
            Q.Next;
        end;

        Fliste.AllSelected:=False;
    end
    else // si pas tout ....
    begin
        InitMove(Fliste.NbSelected,'');

        for i := 0 to Fliste.NbSelected-1 do
        begin
            MoveCur(False);
            Fliste.GotoLeBookMark(i);
            {$IFDEF EAGLCLIENT}
            FEcran.Q.TQ.Seek(Fliste.Row-1);
            {$ENDIF}
            UpdateEnreg(Q);
        end;

        Fliste.ClearSelected;
    end;
    FiniMove;
    FEcran.BChercheClick(nil);
    RestorRow;
end;


function TOF_IMPORTBOB.UpdateEnreg(Q : TQuery) : Boolean;
var
Filename   : string;
rr         : integer;
        procedure UpdateEtat;
        begin
                   ExecuteSQL ('UPDATE YFILESTD SET YFS_BCRIT1="-" WHERE YFS_NOM="'+Q.FindField ('YFS_NOM').asstring+'" AND YFS_CRIT1="'+Q.FindField ('YFS_CRIT1').asstring+'" AND '+
                    'YFS_CRIT2="'+Q.FindField ('YFS_CRIT2').asstring+'" AND ' +
                    'YFS_CRIT3="'+Q.FindField ('YFS_CRIT3').asstring+'" AND ' +
                    'YFS_CRIT4="'+Q.FindField ('YFS_CRIT4').asstring+'" AND ' +
                    'YFS_CRIT5="'+Q.FindField ('YFS_CRIT5').asstring+'" AND ' +
                    'YFS_PREDEFINI="'+Q.FindField ('YFS_PREDEFINI').asstring+'" AND ' +
                    'YFS_NODOSSIER="'+ Numdos +'" AND ' +
                    'YFS_LANGUE="'+Q.FindField ('YFS_LANGUE').asstring+'"');
        end;
begin
       Result := TRUE;
       // Etat n'est pas ouvert
       if Q.FindField ('YFS_BCRIT1').asstring <> 'X' then
            exit;
       // accès interdit
       if Q.FindField ('YFS_BCRIT2').asstring = 'X' then exit;

        if EstSpecif('51502') or (ctxStandard in V_PGI.PGIContexte) then
            Numdos := '000000'
        else
            Numdos := V_PGI.NoDossier;

        Filename := AGL_YFILESTD_GET_PATH(Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring,
        Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring,
        Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, Numdos);

        if FileExists (Filename) then
        begin
                    rr := FileOpen(FileName, fmShareExclusive);
                    if rr = -1 then
                    begin
                         PgiInfo ('Attention le fichier '+ Q.FindField ('YFS_NOM').asstring + ' est encore ouvert');
                         FileClose (rr); Result := FALSE;
                    end
                    else
                    begin
                         FileClose (rr);
                         AGL_YFILESTD_IMPORT(Filename,Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_EXTFILE').asstring,
                            Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring,
                            Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring,'-','-','-','-','-',
                            V_PGI.LanguePrinc,
                            Q.FindField ('YFS_PREDEFINI').asstring,
                            Q.FindField ('YFS_LIBELLE').asstring, Numdos);
                            UpdateEtat;
                            DeleteFile(Filename);
                     end;
        end
        else
                      UpdateEtat;
end;

procedure TOF_IMPORTBOB.OnClose;
var
    Filename : string;
    rr : integer;
begin
  inherited;
{$IFDEF EAGLCLIENT}
    Q := FEcran.Q.TQ;
    Q.Seek(FListe.Row-1);
{$ELSE}
    Q := FEcran.Q;
{$ENDIF}

     if not (Q.Eof) then
     begin
        if ExisteSQL ('SELECT YFS_BCRIT1 from YFILESTD Where  YFS_BCRIT1="X" AND YFS_CODEPRODUIT="'+ Q.FindField ('YFS_CODEPRODUIT').asstring +
                        '" AND YFS_CRIT1="'+Q.FindField ('YFS_CRIT1').asstring+'" AND '+
                        'YFS_PREDEFINI="'+Q.FindField ('YFS_PREDEFINI').asstring+'" AND ' +
                        'YFS_LANGUE="'+Q.FindField ('YFS_LANGUE').asstring+ '" AND ' +
                        'YFS_NODOSSIER="'+ Numdos +'"') then
        begin
            case PGIAskCancel  ('Attention, il existe des documents ouverts, souhaitez-vous les sauvegarder dans la base ?') of
            mrNo     : begin
                       Q.First;
                       while (Not Q.EOF) do
                       begin

                           if Q.FindField ('YFS_BCRIT1').asstring ='X' then
                           begin
                                Filename := GetWindowsTempPath +'PGI\STD\'+Q.FindField ('YFS_CODEPRODUIT').asstring+'\'+Numdos+'\';
                                if Q.FindField ('YFS_CRIT1').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT1').asstring+'\';
                                if Q.FindField ('YFS_CRIT2').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT2').asstring+'\';
                                if Q.FindField ('YFS_CRIT3').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT3').asstring+'\';
                                if Q.FindField ('YFS_CRIT4').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT4').asstring+'\';
                                if Q.FindField ('YFS_CRIT5').asstring <> '' then Filename := Filename + Q.FindField ('YFS_CRIT5').asstring+'\';
                                Filename := Filename + Q.FindField ('YFS_LANGUE').asstring+'\'+Q.FindField ('YFS_PREDEFINI').asstring+'\'+
                                Q.FindField ('YFS_NOM').asstring;
                                if FileExists (Filename) then
                                begin
                                    rr := FileOpen(FileName, fmShareExclusive);
                                    if rr = -1 then
                                    begin
                                     PgiInfo ('Attention le fichier '+ Q.FindField ('YFS_NOM').asstring + ' est encore ouvert');
                                     Ecran.ModalResult := mrNone;  LastError := 4;
                                     FileClose (rr);
                                     exit;
                                    end;
                                    FileClose (rr);
                                    DeleteFile(Filename);
                                end;
                            end;
                            Q.Next;
                      end;
                       ExecuteSQL ('UPDATE YFILESTD SET YFS_BCRIT1="-" Where YFS_NODOSSIER="'+Numdos+'" AND YFS_CRIT1="'+SType+'"');
            end;
            mrYes    :
            begin
                Q.First;
                while (Not Q.EOF) do
                begin
                     if Q.FindField ('YFS_BCRIT1').asstring ='X' then
                        if not UpdateEnreg(Q ) then
                        begin
                                     Ecran.ModalResult := mrNone; LastError := 4;
                                     exit;
                        end;
                    Q.Next;
                end;
            end;
             mrCancel :
             begin
                Ecran.ModalResult := mrNone;   LastError := 4;
                exit;
             end;
            end;
        end;
    end;
    TheTob.free;
end;

procedure TOF_IMPORTBOB.SaveRow;
begin
{$IFDEF EAGLCLIENT}
   Q := FEcran.Q.TQ;
   Q.Seek(FListe.Row-1);
   wBookMark := FListe.Row;
{$ELSE}
   Q := FEcran.Q;
   wBookMark := FListe.DataSource.DataSet.GetBookmark;
{$ENDIF}
end;

procedure TOF_IMPORTBOB.RestorRow;
begin
{$IFDEF EAGLCLIENT}
      Q := FEcran.Q.TQ;
      FListe.Row := wBookMark;
      if FListe.Row >0 then
        Q.Seek( FListe.Row )
      else
        Q.First;
{$ELSE}
     Q.GotoBookmark(wBookMark);
{$ENDIF}
end;

procedure TOF_IMPORTBOB.DupliqueOnClick(Sender: TObject);
var
CodeRetour : integer;
FName,Ext  : string;
ind, i     : integer;
Filename   : string;
FichierAIntegrer : string;
WWW              : string;
begin
   SaveRow;
   if Q.EOF then
   begin
      PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
   end;

   Ext := ExtractFileName (Q.FindField ('YFS_NOM').asstring);
   Fname   := ReadTokenPipe (Ext, '.');
   Ext := '.'+Ext;
   Ind := 1;
   i :=  Pos ('_', Fname);
   if i > 0 then
   Fname := copy(Fname, 0, i-1);

   While (ExisteSql ('SELECT * FROM YFILESTD WHERE YFS_NOM="'+Fname+'_'+ IntTostr(Ind)+Ext+'" AND YFS_CRIT1="'+SType+
   '" AND YFS_PREDEFINI="'+THValComboBox(GetControl('YFS_PREDEFINI')).Value+'" AND YFS_NODOSSIER="'+ Numdos+'"')) do inc (Ind);
   Fname   := Fname+ '_'+ IntTostr(Ind)+Ext;

   if Q.EOF then
   begin
      PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
   end;

  if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then
  FichierAIntegrer :=  V_PGI.DosPath +'\'+ Q.FindField ('YFS_NOM').asstring
  else
  FichierAIntegrer :=  V_PGI.DosPath +Q.FindField ('YFS_NOM').asstring;

  Filename := AGL_YFILESTD_GET_PATH(Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring,
  Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring,
  Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, Numdos);

  if not FileExists (Filename) then
  begin
        CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring
        ,Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, Q.FindField ('YFS_NODOSSIER').asstring);
        if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier')
        else
            Filename := FichierAIntegrer; // fiche 22123 base multicsoc en PGI
  end
  else
        CodeRetour := -1;

   if CodeRetour = -1 then
   begin
        Fname := ExtractFileDir(Filename)+'\'+ Fname;
        Copyfile (PChar( Filename), Pchar(Fname),TRUE);

       if EstSpecif('51502') then
          WWW := 'CEG'
       else
       begin
            if (ctxStandard in V_PGI.PGIContexte) then
                WWW := 'STD'
            else
                WWW := 'DOS';
       end;

        AGL_YFILESTD_IMPORT(Fname,Q.FindField ('YFS_CODEPRODUIT').asstring, ExtractFileName(Fname), Q.FindField ('YFS_EXTFILE').asstring,
                      Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring,
                      Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring,'-','-','-','-','-',
                      V_PGI.LanguePrinc, WWW,
                      Copy(Q.FindField ('YFS_LIBELLE').asstring,1, 30) + '.'+IntTostr(Ind), Numdos);
     FEcran.BChercheClick(nil) ;
     RestorRow;
     DeleteFile(Fname);
     DeleteFile(Filename);

   end;
end;

procedure TOF_IMPORTBOB.PredefiniChange(Sender: TObject);
begin
if (THValComboBox (GetControl ('YFS_PREDEFINI')).value = 'DOS') then
       THEdit(GetControl('XX_WHERE')).Text := 'YFS_CRIT1="' + SType + '" AND' +
       ' YFS_NODOSSIER="'+ Numdos +'"'
else // fiche 20449
if ((THValComboBox (GetControl ('YFS_PREDEFINI')).value = 'CEG')) or
((THValComboBox (GetControl ('YFS_PREDEFINI')).value = 'STD')) then
       THEdit(GetControl('XX_WHERE')).Text := 'YFS_CRIT1="' + SType + '" AND' +
       ' YFS_NODOSSIER="000000"'
else
       THEdit(GetControl('XX_WHERE')).Text := 'YFS_CRIT1="' + SType + '" AND' +
       ' ((YFS_NODOSSIER="000000") or (YFS_NODOSSIER="'+ Numdos +'"))'

end;

{$IFDEF SCANGED}
procedure TOF_IMPORTBOB.GedOnClick(Sender: TObject);
var
Q1            : TQuery;
FileGuid      : string;
TImp,vTobRech : TOB;
SQL,Numdos    : string;
Filename      : string;
CodeRetour    : integer;
Libelle       : string;
begin
      if (ctxPCL in V_PGI.PGIContexte) then
      begin
       if PgiAsk ('Confirmez-vous l''archivage dans le dossier client ?') <> mrYes then exit;

          if EstSpecif('51502') then
                Numdos := '000000'
          else
                Numdos := V_PGI.NoDossier;
          SaveRow;
          if Q.EOF then
          begin
            PGIInfo ('Aucune feuille sélectionnée'); RestorRow; exit;
          end;

          Libelle := Q.FindField ('YFS_LIBELLE').asstring;

          SQL := 'SELECT YFS_FILEGUID FROM YFILESTD Where YFS_CODEPRODUIT="COMPTA" AND '+
        ' YFS_CRIT1="'+SType+'" AND'+
        ' YFS_NOM="'+Q.FindField ('YFS_NOM').asstring+'" AND'+
        ' YFS_LANGUE="'+V_PGI.LanguePrinc+'" AND'+
        ' YFS_NODOSSIER="'+Numdos+'" AND'+
        ' YFS_PREDEFINI="'+THValComboBox (GetControl ('YFS_PREDEFINI')).Value+'"';
            if Q.FindField ('YFS_BCRIT1').asstring = 'X' then begin PGIInfo ('Document extrait'); exit; end;

            Filename := AGL_YFILESTD_GET_PATH(Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring,
            Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring,
            Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, Numdos);

            if not FileExists (Filename) then
            begin
                  CodeRetour :=  AGL_YFILESTD_EXTRACT (Filename, Q.FindField ('YFS_CODEPRODUIT').asstring, Q.FindField ('YFS_NOM').asstring, Q.FindField ('YFS_CRIT1').asstring, Q.FindField ('YFS_CRIT2').asstring, Q.FindField ('YFS_CRIT3').asstring
                  ,Q.FindField ('YFS_CRIT4').asstring, Q.FindField ('YFS_CRIT5').asstring, FALSE, Q.FindField ('YFS_LANGUE').asstring, Q.FindField ('YFS_PREDEFINI').asstring, Numdos);
                  if CodeRetour <> -1 then PGIInfo ('Erreur d''accès au fichier')
            end;


            Q1 := OpenSQL ('SELECT * FROM NFILES WHERE NFI_FILEGUID IN ('+SQL+')', TRUE);
            if not Q1.EOF then
            begin
                    vTobRech := TOB.Create( 'NFILES' , nil , -1 ) ;
                    vTobRech.SelectDB( '', Q1 ) ;
                    Ferme (Q1);

                    FileGuid := vTobRech.Getvalue('NFI_FILEGUID');
                    TImp :=TOB.Create('YFILES',Nil,-1) ;
                    TImp.PutValue ('YFI_FILEID', vTobRech.Getvalue('NFI_FILEID'));
                    TImp.PutValue ('YFI_FILEGUID', vTobRech.Getvalue('NFI_FILEGUID'));
                    TImp.PutValue ('YFI_FILENAME', vTobRech.Getvalue('NFI_FILENAME'));
                    TImp.PutValue ('YFI_FILESIZE', vTobRech.Getvalue('NFI_FILESIZE'));
                    TImp.PutValue ('YFI_FILECOMPRESSED', vTobRech.Getvalue('NFI_FILECOMPRESSED'));
                    TImp.PutValue ('YFI_FILESTORAGE', vTobRech.Getvalue ('NFI_FILESTORAGE'));
                    TImp.PutValue ('YFI_CREATEUR', vTobRech.Getvalue('NFI_CREATEUR'));
                    TImp.InsertOrUpdateDB(True);
                    TImp.free;
                    vTobRech.free;
            end else Ferme (Q1);
            Q1 := OpenSQL ('SELECT * FROM NFILEPARTS WHERE NFS_FILEGUID IN ('+SQL+')', TRUE);
            if not Q1.EOF then
            begin
                    vTobRech := TOB.Create( 'NFILEPARTS' , nil , -1 ) ;
                    vTobRech.SelectDB( '', Q1 ) ;
                    Ferme (Q1);
                    TImp :=TOB.Create('YFILEPARTS',Nil,-1) ;
                    TImp.PutValue ('YFP_FILEID',  vTobRech.Getvalue('NFS_FILEID'));
                    TImp.PutValue ('YFP_FILEGUID', vTobRech.Getvalue ('NFS_FILEGUID'));
                    TImp.PutValue ('YFP_PARTID', vTobRech.Getvalue('NFS_PARTID'));
                    TImp.PutValue ('YFP_DATA', vTobRech.Getvalue ('NFS_DATA'));
                    TImp.InsertOrUpdateDB(True);
                    TImp.free;
                    vTobRech.free;
            end else Ferme (Q1);
            if FileGuid <> '' then
            begin
                MyImportGed (FileGuid, Libelle);
            end;
            DeleteFile(Filename);
      end;
end;
{$ENDIF}

initialization
  registerclasses([TOF_IMPORTBOB]);

end.

