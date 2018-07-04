{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 11/06/2001
Modifié le ... :   /  /
Description .. : Fiche 4 Jpegs 4 Mémos (LiensOLE) - Passer GA_ARTICLE en argument
Mots clefs ... : JPG;JPEG;MEMO
*****************************************************************}
Unit UTofImgsMemos;

interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF, entgc ; 

type
  TOF_ImagesMemosX4 = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S : String); override;
    procedure OnClose; override;

    procedure OnInsClick(Sender : TObject);
    procedure OnDelClick(Sender : TObject);
    procedure OnEditChange(Sender : TObject);

  private
    GaArticle : String;
    FileName, EmploiJpg, EmploiMem : Array[1..4] of string;
    JModified, MModified : Array[1..4] of boolean;

    procedure GetParamz;
    procedure LoadLiensOLE;
    procedure SaveLiensOLE;

  end;

{$IFDEF CCS3}
Const nbImages = 1 ;
{$ELSE}
Const nbImages = 1 ;   // pour les tests normalement 4
{$ENDIF}

implementation
uses jpeg, extctrls, utilPGI, Htb97, LookUp, ParamSoc, HRichOLE;

procedure TOF_ImagesMemosX4.OnNew;
begin
  inherited;
end ;

procedure TOF_ImagesMemosX4.OnDelete;
begin
  inherited;
end ;

procedure TOF_ImagesMemosX4.OnUpdate;
begin
  inherited;
  Transactions(SaveLiensOLE, 3);
end ;

procedure TOF_ImagesMemosX4.OnLoad;
begin
  inherited;
  LoadLiensOLE;
end ;

procedure TOF_ImagesMemosX4.OnArgument(S : String);
var i : integer;
begin
  inherited;
  GaArticle := S;

  GetParamz;

  for i := 1 to nbImages do
  begin
    TToolbarButton97(GetControl('BIMAGE_INSERT'+inttostr(i))).OnClick := OnInsClick;
{$IFNDEF CCS3}
    TToolbarButton97(GetControl('IMAGE'+inttostr(i))).OnClick := OnInsClick;
{$ENDIF}
    TToolbarButton97(GetControl('BIMAGE_DELETE'+inttostr(i))).OnClick := OnDelClick;
//    THEdit(GetControl('TETITRE'+inttostr(i))).OnChange := OnEditChange;
//    THEdit(GetControl('TEMTITRE'+inttostr(i))).OnChange := OnEditChange;
//    THRichEditOLE(GetControl('MEMO'+inttostr(i))).OnChange := OnEditChange;
    JModified[i] := false;
    MModified[i] := false;
  end;
end ;

procedure TOF_ImagesMemosX4.OnClose;
begin
  inherited;
end;

procedure TOF_ImagesMemosX4.OnInsClick(Sender : TObject);
var n : string;
begin
    n := Copy(TControl(Sender).Name, Length(TControl(Sender).Name), 1);
    if GetFileName(tfOpenBMP, '*.JPG;*.JPEG', FileName[strtoint(n)]) then
    begin
      TImage(GetControl('IMAGE'+n)).Picture.LoadFromFile(FileName[strtoint(n)]);
      JModified[strtoint(n)] := true;
      SetControlText('ETITRE'+n, ExtractFileName(FileName[strtoint(n)]));
      SetFocusControl('ETITRE'+n);
    end;
end;

procedure TOF_ImagesMemosX4.OnDelClick(Sender : TObject);
var n : string;
begin
    n := Copy(TControl(Sender).Name, Length(TControl(Sender).Name), 1);
    TImage(GetControl('IMAGE'+n)).Picture := nil;
    SetControlText('ETITRE'+n, '');
    JModified[strtoint(n)] := true;
end;

procedure TOF_ImagesMemosX4.OnEditChange(Sender : TObject);
var n : integer;
begin
    if not TCustomEdit(Sender).Modified then exit;

    n := strtoint(Copy(TControl(Sender).Name, Length(TControl(Sender).Name), 1));
    if UpperCase(Copy(TControl(Sender).Name, 1, 7)) = 'TETITRE' then JModified[n] := true
                                                                else MModified[n] := true;
end;

procedure TOF_ImagesMemosX4.GetParamz;
var i : integer;
begin
{$IFDEF CCS3}
      EmploiJpg[1] := VH_GC.GCPHOTOFICHE;
      SetControlText('TETITRE1', RechDom('GCEMPLOIBLOB', EmploiJpg[1], false));
{$ELSE}
   for i := 1 to nbImages do
   begin
      EmploiJpg[i] := GetParamSoc('SO_GCCDJPG'+inttostr(i));
      SetControlText('TETITRE'+inttostr(i), RechDom('GCEMPLOIBLOB', EmploiJpg[i], false));
      EmploiMem[i] := GetParamSoc('SO_GCCDMEM'+inttostr(i));
      SetControlText('TEMTITRE'+inttostr(i), RechDom('GCEMPLOIBLOB', EmploiMem[i], false));
   end;
{$ENDIF}
end;

procedure TOF_ImagesMemosX4.LoadLiensOLE;
var i : integer;
    Q : TQuery;
begin
   for i := 1 to nbImages do
   begin
      Q := OpenSQL('SELECT LO_LIBELLE, LO_OBJET FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_QUALIFIANTBLOB="PHJ" AND LO_EMPLOIBLOB="'+EmploiJpg[i]+'" AND LO_IDENTIFIANT="'+GaArticle+'"', true);
      if not Q.EOF then begin try LoadBitmapFromChamp(Q, 'LO_OBJET', TImage(GetControl('IMAGE'+inttostr(i))), true);
                              except TImage(GetControl('IMAGE'+inttostr(i))).Picture := nil; end;
                              SetControlText('ETITRE'+inttostr(i), Q.FieldByName('LO_LIBELLE').AsString); end
                   else begin TImage(GetControl('IMAGE'+inttostr(i))).Picture := nil; //PGIInfo('Pas d''image', '');
                              SetControlText('ETITRE'+inttostr(i), ''); end;
      Ferme(Q);
      Q := OpenSQL('SELECT LO_LIBELLE, LO_OBJET FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_QUALIFIANTBLOB="MEM" AND LO_EMPLOIBLOB="'+EmploiMem[i]+'" AND LO_IDENTIFIANT="'+GaArticle+'"', true);
      if not Q.EOF then begin SetControlText('MEMO'+inttostr(i), Q.FieldByName('LO_OBJET').AsString);
                              SetControlText('EMTITRE'+inttostr(i), Q.FieldByName('LO_LIBELLE').AsString); end
                   else begin SetControlText('MEMO'+inttostr(i), '');
                              SetControlText('EMTITRE'+inttostr(i), ''); end;
      Ferme(Q);
   end;
end;

procedure TOF_ImagesMemosX4.SaveLiensOLE;
var i, rb : integer;
    Q, QQ : TQuery;
//    Sjpg : TMemoryStream;
begin
   for i := 1 to nbImages do
   begin
     if JModified[i] then
      if (TImage(GetControl('IMAGE'+inttostr(i))).Picture.Graphic <> nil) and
         (not TImage(GetControl('IMAGE'+inttostr(i))).Picture.Graphic.Empty) then
      begin
        Q := OpenSQL('SELECT * FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_QUALIFIANTBLOB="PHJ" AND LO_EMPLOIBLOB="'+EmploiJpg[i]+'" AND LO_IDENTIFIANT="'+GaArticle+'"', false);
        Q.Edit;

        if Q.State = dsInsert then
        begin
          Q.FieldByName('LO_TABLEBLOB').AsString := 'GA';
          Q.FieldByName('LO_QUALIFIANTBLOB').AsString := 'PHJ';
          Q.FieldByName('LO_EMPLOIBLOB').AsString := EmploiJpg[i];
          Q.FieldByName('LO_IDENTIFIANT').AsString := GaArticle;
          QQ := OpenSQL('SELECT MAX(LO_RANGBLOB) AS MAXRANG FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_IDENTIFIANT="'+GaArticle+'"', true);
          if not QQ.EOF then rb := QQ.FieldByName('MAXRANG').AsInteger+1
                        else rb := 1;
          Ferme(QQ);
          Q.FieldByName('LO_RANGBLOB').AsInteger := rb;
          Q.FieldByName('LO_PRIVE').AsString := '-';
        end;
  //      Q.FieldByName('LO_DATEBLOB').AsDateTime := =DATECREATION,D
        Q.FieldByName('LO_LIBELLE').AsString := GetControlText('ETITRE'+inttostr(i));

        TBlobField(Q.FieldByName('LO_OBJET')).LoadFromFile(FileName[i]);

        Q.Post;
        Ferme(Q);
      end
       else ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_QUALIFIANTBLOB="PHJ" AND LO_EMPLOIBLOB="'+EmploiJpg[i]+'" AND LO_IDENTIFIANT="'+GaArticle+'"');

//////  Memos
{$IFNDEF CCS3}

//     if MModified[i] then
//     begin
        Q := OpenSQL('SELECT * FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_QUALIFIANTBLOB="MEM" AND LO_EMPLOIBLOB="'+EmploiMem[i]+'" AND LO_IDENTIFIANT="'+GaArticle+'"', false);
        Q.Edit;

        if Q.State = dsInsert then
        begin
          Q.FieldByName('LO_TABLEBLOB').AsString := 'GA';
          Q.FieldByName('LO_QUALIFIANTBLOB').AsString := 'MEM';
          Q.FieldByName('LO_EMPLOIBLOB').AsString := EmploiMem[i];
          Q.FieldByName('LO_IDENTIFIANT').AsString := GaArticle;
          QQ := OpenSQL('SELECT MAX(LO_RANGBLOB) AS MAXRANG FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_IDENTIFIANT="'+GaArticle+'"', true);
          if not QQ.EOF then rb := QQ.FieldByName('MAXRANG').AsInteger+1
                        else rb := 1;
          Ferme(QQ);
          Q.FieldByName('LO_RANGBLOB').AsInteger := rb;
          Q.FieldByName('LO_PRIVE').AsString := '-';
        end;
        Q.FieldByName('LO_LIBELLE').AsString := GetControlText('EMTITRE'+inttostr(i));
        Q.FieldByName('LO_OBJET').AsString := GetControlText('MEMO'+inttostr(i));

        Q.Post;
        Ferme(Q);
//     end;
{$ENDIF}
  end;
end;

initialization
RegisterClasses([TOF_ImagesMemosX4]);
end.

