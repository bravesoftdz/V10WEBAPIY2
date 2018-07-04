unit UTOFMULINFOCA3;

interface

uses  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,ComObj,db,dbTables,
      ComCtrls,ExtCtrls, Menus, OleCtnrs, StdCtrls,inifiles, ShellApi,Mul,
      FE_Main,HCtrls,HEnt1,HDB,HMsgBox,DBCtrls, HRichEdt, HRichOLE, UTob,UTOF,Vierge,HTB97,UiUtil,
      Lia_Commun,HStatus,UTOM,Ent1,SaisUtil,PGIENV,ParamSoc,dbgrids,Grids,LettUtil,
      M3FP,HPanel,Lia_edition,LIA_FCT_EXTERNE,InitFichCa3;

Const slach='/';
Type
  TOF_INFOCA3=Class(TOF)
   procedure OnArgument(Arguments:String);override;
   procedure OnLoad ; override ;
   procedure OnUpdate ; override ;
   private
   NumDecla                    : integer;
   TOB_VALRUB,TOB_TEMPO        :TOB;
   ChargeValrub                : string;
   DateEncours                 : TDateTime;
   procedure BTCreatClick (Sender : TObject);
   procedure DessineCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
   procedure OnDrawColumnCellFListe(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
   FUNCTION QUELEXODTCA3(DD : TDateTime) : String ;

end;
const PREFIXE_VAL='F10';
const PREFIXE_STRUC='D10';
procedure CA3PrintDeclaTVA (Vr : Boolean);

implementation

procedure TOF_INFOCA3.OnArgument(Arguments:String);
begin
NumDecla := 1;
ChargeValrub := ReadTokenSt (Arguments);
TToolbarButton97(GetControl('BINSERT')).Onclick := BTCreatClick;
THDBGrid (GetControl('FListe')).OnDrawColumnCell := OnDrawColumnCellFListe;
end;

procedure TOF_INFOCA3.OnUpdate;
BEGIN
TFMul(Ecran).Q.Locate ('CA3_DATEDEBUT', LiasseEnv.DATEENCOURS, []);
END;

procedure TOF_INFOCA3.DessineCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
Var Value : string ;
FListe: THDBGrid;
begin
Inherited ;
FListe := THDBGrid (GetControl('FListe'));
{
if (Value=DateToStr(LiasseEnv.DATEENCOURS)) then
begin
      THDBGrid (GetControl('FListe')).Canvas.Font.Color:= clRed;
      THDBGrid (GetControl('FListe')).Canvas.TextOut ( (Rect.Left+Rect.Right) div 2 - THDBGrid (GetControl('FListe')).Canvas.TextWidth(Value) div 2,
      (Rect.Top+ Rect.Bottom) div 2 - THDBGrid (GetControl('FListe')).Canvas.TextHeight(Value) div 2, Value);
end;
}
If Assigned(Column.Field) and (Column.Field is TDateTimeField) then
   begin
   if (TDateTimeField(Column.Field).AsDateTime=DATEENCOURS) then
   begin
   Value := Column.Field.DisplayText;
      THDBGrid (GetControl('FListe')).Canvas.Font.Color:= clRed;
//      THDBGrid (GetControl('FListe')).Canvas.TextOut ( 13,
      THDBGrid (GetControl('FListe')).Canvas.TextOut ( ((Rect.Left+Rect.Right) div 2 - THDBGrid (GetControl('FListe')).Canvas.TextWidth(Value) div 2)-18,
      (Rect.Top+ Rect.Bottom) div 2 - THDBGrid (GetControl('FListe')).Canvas.TextHeight(Value) div 2, Value);
   end;

   if (TDateTimeField(Column.Field).EditMask='n') and
      ((TDateTimeField(Column.Field).IsNull) or (TDateTimeField(Column.Field).AsDateTime=iDate1900)) then
      BEGIN
      FListe.Canvas.FillRect(Rect);
      END ;

   end ;

If Assigned(Column.Field) and (Column.Field is TMemoField) then
   begin
   if TMemoField(Column.Field).IsNull then Value:='' else
      BEGIN
      if TRUE or (TFMul(Ecran).MemoStyle=msFirstline) or (V_PGI.Driver=dbMSACCESS) then
         BEGIN
         TFMul(Ecran).RR.Lines.Assign(TMemoField(Column.Field)) ;
         if TFMul(Ecran).RR.Lines.Count>0 then Value:=TFMul(Ecran).RR.Lines[0] else Value:='' ;
         END else  Value:='r' ;
      END ;
   FListe.Canvas.FillRect(Rect);
   Case TFMul(Ecran).MemoStyle of
      msNone : ;
      msFirstLine : FListe.Canvas.TextOut(Rect.Left, Rect.Top, Value);
      msCheckBox,msCoche,msBook :
                  BEGIN
                  if TFMul(Ecran).MemoStyle=msCheckBox then BEGIN if Value<>'' then Value:= 'þ' Else Value:='¨' ;END ;
                  if TFMul(Ecran).MemoStyle=msCoche then BEGIN if Value<>'' then Value:= 'ü' Else Value:=' ' ; END ;
                  if TFMul(Ecran).MemoStyle=msBook then BEGIN if Value<>'' then Value:= '&' Else Value:=' ' ;END ;
                  FListe.Canvas.Font.Name:='Wingdings' ;
                  FListe.Canvas.Font.Size:=10 ;
                  FListe.Canvas.TextOut ( (Rect.Left+Rect.Right) div 2 - FListe.Canvas.TextWidth(Value) div 2,
                                        (Rect.Top+ Rect.Bottom) div 2 - FListe.Canvas.TextHeight(Value) div 2, Value);
                  END ;
      end;
   end ;
if (TFMul(Ecran).CheckBoxStyle<>csBoolean) and Assigned(Column.Field) and (Column.Field.DataType=ftString) and (Column.Field.Size=1) then
   BEGIN
   Value := Column.Field.DisplayText;

   if ((Value='X') or (Value='-')) then
      BEGIN
      Case TFMul(Ecran).CheckBoxStyle of
         csCheckBox : if Value='X' then Value:= 'þ' Else Value:='¨' ;
         csCoche : if Value='X' then Value:= 'ü' Else Value:=' ' ;
         END ;
      FListe.Canvas.FillRect(Rect);
      FListe.Canvas.Font.Name:='Wingdings' ;
      FListe.Canvas.Font.Size:=10 ;
      FListe.Canvas.TextOut ( (Rect.Left+Rect.Right) div 2 - FListe.Canvas.TextWidth(Value) div 2,
      (Rect.Top+ Rect.Bottom) div 2 - FListe.Canvas.TextHeight(Value) div 2, Value);
      END ;
   END ;

end;

procedure TOF_INFOCA3.OnDrawColumnCellFListe(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
FListe: THDBGrid;
begin
  FListe := THDBGrid (GetControl('FListe'));
If Assigned(Column.Field) and (Column.Field is TDateTimeField) then
begin
  if not (State >= [gdSelected]) then
  begin
    if (TDateTimeField(Column.Field).AsDateTime=DATEENCOURS) then
    begin
      if V_PGI.NumAltCol = 4 then // Bleu
        (Sender as TDBGrid).Canvas.Brush.Color := ($00B5FAFD) // Jaune
      else
        (Sender as TDBGrid).Canvas.Brush.Color := ($00F3E8D1); // Bleu
    end;
  FListe.DefaultDrawColumnCell(Rect,DataCol,Column,State);
  end;
end;
end;

procedure TOF_INFOCA3.OnLoad;
var
Q1                           : TQuery;
Codeperfis                   : integer;
begin
DateEncours := iDate1900;
Q1 := OpenSql ('SELECT CP3_DATEENCOURS FROM CA3_PARAMETRE WHERE CP3_CODEPER='+ LiasseEnv.CODEPER, TRUE);
if not Q1.EOF then
begin
     DateEncours := Q1.FindField ('CP3_DATEENCOURS').Asdatetime;
     If (Ecran <> nil)  then Ecran.Caption:= 'Liste des CA3 disponibles : ' + DateToStr(DateEncours);
     UpDateCaption(Ecran);
end;
Ferme (Q1);
end;

procedure TOF_INFOCA3.BTCreatClick (Sender : TObject);
var
Tobca3                         : TOB;
Q1                             : TQuery;
Periodicite                    : string;
Datedeb,Datefin,Dateech        : TDateTime;
Jourech                        : integer;
Date,Date1                     : string;
Result                         : Boolean;
Exercice                       : string;
Redevabilite                   : string;
a, m, j                        : word ;
Datedeb2                       : TDateTime;
begin
     DecodeDate(LiasseEnv.DateEncours,a,m,j) ;
     if a  < StrToint (LiasseEnv.MILLESIME) then
     begin
          DecodeDate(VH^.Encours.fin,a,m,j) ;
          if a  < StrToint (LiasseEnv.MILLESIME) then
          begin
             PGIINFO ('La date de votre déclaration ne correspond pas à Millésime ' + LiasseEnv.MILLESIME ,'Satut des déclarations CA3');  exit;
          end
          else
          begin
                  DecodeDate(LiasseEnv.DateEncours,a,m,j) ;
                  Datedeb2 := LiasseEnv.DateEncours;
                  while a  < StrToint (LiasseEnv.MILLESIME) do
                  begin
                       Datedeb := PlusMois (Datedeb2, 1);
                       DecodeDate(Datedeb,a,m,j) ;
                       Datedeb2 := Datedeb;
                  end;
                  LiasseEnv.DateEncours := Datedeb2;

          end;
     end;

     Result:=ExisteSQL('SELECT CA3_STATUT from CA3_DECLARATION Where CA3_STATUT="-"');
     if Result then
     begin
             PGIINFO ('Il existe des déclarations CA3 ouvertes','Satut des déclarations CA3');  exit;
     end;
     Datedeb := iDate1900;
     Datefin := iDate1900;
     TObca3 := TOB.Create('CA3_DECLARATION',Nil,-1);

      Q1 := Opensql ('SELECT CP3_PERIODICITE,CP3_JOURDECLA,CP3_REDEVABILITE FROM CA3_PARAMETRE where CP3_CODEPER='+LiasseEnv.CODEPER, TRUE);
      if not Q1.EOF then
      begin
      Periodicite := Q1.FindField ('CP3_PERIODICITE').asstring;
      Jourech := Q1.FindField ('CP3_JOURDECLA').asinteger;
      Redevabilite := Q1.FindField ('CP3_REDEVABILITE').asstring;
      end;
      ferme (Q1);

     if Periodicite = '' then
     begin
             PGIINFO ('Périodicité non renseigné','Information sur la période');  exit;
     end;
     if Jourech = 0 then
     begin
             PGIINFO ('Jour dépôt déclaration non renseigné','Jour dépôt déclaration');  exit;
     end;
     Q1 := Opensql ('SELECT CA3_DATEFIN FROM CA3_DECLARATION ORDER BY CA3_DATEDEBUT DESC', TRUE);
     if Q1.EOF then Datedeb := LiasseEnv.DateEncours
     else
     begin
          Datefin := Q1.FindField ('CA3_DATEFIN').asdatetime;
          Datedeb := Datefin;
          DecodeDate (Datedeb, a, m, j);
          if J<>1 then Datedeb := Datedeb+1;
     end;
     Ferme (Q1);

     if Datedeb = iDate1900 then
     begin
             PGIINFO ('date début d''exercice non renseigné','Date début d''exercice');  exit;
     end;

     if Periodicite = 'MEN' then
        Datefin := PlusMois (Datedeb, 1);

     if Periodicite = 'BIM' then
        Datefin := PlusMois (Datedeb, 2);

     if Periodicite = 'TRI' then
        Datefin := PlusMois (Datedeb, 3);

     if Periodicite = 'SEM' then
        Datefin := PlusMois (Datedeb, 6);

     if Periodicite = 'ANN' then
        Datefin := VH^.Encours.fin;

     DecodeDate (Datedeb, a, m, j);
     if J=1 then Datefin := Datefin-1;

     Dateech := Datefin;
     Dateech :=  StrToDate(IntToStr (Jourech) + copy (DateToStr(Dateech), 3, length (DateToStr(Dateech))));
     Dateech := PlusMois (Dateech, 1);
     
     Exercice := QUELEXODTCA3  (Datedeb);
     if Exercice = '' then
     begin
          PGIInfo ('Exercice non ouvert', 'Exercice'); exit;
     end;
     TObca3.AddChampSupValeur('CA3_EXERCICE',Exercice);

     TObca3.AddChampSupValeur('CA3_LIBELLE','Déclaration du :' + DateToStr (Datedeb));
     TObca3.AddChampSupValeur('CA3_DATEDEBUT',Datedeb);
     TObca3.AddChampSupValeur('CA3_DATEFIN',Datefin);
     TObca3.AddChampSupValeur('CA3_STATUT','-');
     TObca3.AddChampSupValeur('CA3_DATEECH',Dateech);
     TObca3.AddChampSupValeur('CA3_PERIODICITE',Periodicite);
     DateEncours := Datedeb;
     if (TObca3.ExistDB) then  // maj ca3_declaration
          TObca3.UpdateDB (FALSE)
     else
          TObca3.InsertOrUpdateDB(FALSE);
     if TObca3<>nil then TObca3.free;

     Result := ExisteSQL ('SELECT CP3_DATEENCOURS FROM CA3_PARAMETRE WHERE CP3_CODEPER='+ LiasseEnv.Codeper);
        Date :=  DateToStr(DateEncours);
        Date1 := ReadTokenPipe(Date,slach);
        Date1 := Redevabilite;
        Date1 := Date1+ReadTokenPipe(Date,slach);
        Date1 := Date1+copy (Date, 3, 2);

     if Result then
     begin
           ExecuteSQL ('Update CA3_PARAMETRE set CP3_DATEENCOURS ="' + USDATETIME(DateEncours)+'",'
           + 'CP3_FPERIODE ="' + Date1 + '"');
           LiasseEnv.DateEncours := DateEncours;
     end
     else
     begin
             ExecuteSql('INSERT INTO CA3_PARAMETRE (CP3_CODEPER,CP3_DATEENCOURS,CP3_FPERIODE)' +
                        'Values ('+LiasseEnv.CODEPER+',"'+USDATETIME (DateEncours)+'"'+ Date1+'")');
             LiasseEnv.DateEncours := VH^.Encours.Deb;
     end;
     If (Ecran <> nil)  then Ecran.Caption:= 'Liste des CA3 disponibles : ' + DateToStr(DateEncours);
     TFMul(Ecran).ChercheClick ;
     UpDateCaption(Ecran);
end;

// la suppression
Type
  TEvtSup = Class
    Exercice    : string;
    Datedebut   : string;
    procedure Degageca3;
end;
var  EvtSup : TEvtSup;

procedure TEvtSup.Degageca3;
begin
      if ExecuteSQL('DELETE FROM CA3_DECLARATION WHERE CA3_EXERCICE="'+Exercice+'"'
      + ' AND CA3_DATEDEBUT="'+ Datedebut+'"')<>1 then V_PGI.IoError:=oeUnknown ;
end;

procedure SupprimeListeEnregca3(L : THDBGrid; Q :TQuery; St:string);
var i,j     : integer;
Texte       : string;
Q1          : TQuery;
Date,Date1  : string;
begin

  if (L.NbSelected=0) and (not L.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;
  Texte:='Vous allez supprimer définitivement les informations.#13#10Confirmez vous l''opération ?';
  if HShowMessage('0;Suppression;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;
  if L.AllSelected then
  BEGIN
    Q.First;
    while Not Q.EOF do
    BEGIN
      MoveCur(False);
      EvtSup := TEvtSup.Create;
      EvtSup.exercice := Q.FindField('CA3_EXERCICE').AsString;
      EvtSup.Datedebut := USDATETIME(Q.FindField('CA3_DATEDEBUT').AsDateTime);

       if Transactions(EvtSup.Degageca3,5)<>oeOK then
       BEGIN
       MessageAlerte('Suppression impossible') ; exit;
       END;
       ExecuteSQL('DELETE FROM F3310_VALRUB WHERE F10_MILLESIME="'+LiasseEnv.MILLESIME+'"'
             + ' AND F10_NUMAUTO=1 AND F10_DATEEFFET="'+ EvtSup.Datedebut+'"');
      EvtSup.Free;
      Q.Next;
    END;
    L.AllSelected:=False;
    LiasseEnv.DateEncours := VH^.Encours.Deb;
  END
  ELSE
  BEGIN
    InitMove(L.NbSelected,'');
    for i:=0 to L.NbSelected-1 do
    BEGIN
      MoveCur(False);
      EvtSup := TEvtSup.Create;
      EvtSup.exercice := Q.FindField('CA3_EXERCICE').AsString;
      EvtSup.Datedebut     := USDATETIME(Q.FindField('CA3_DATEDEBUT').AsDateTime);

       if Transactions(EvtSup.Degageca3,5)<>oeOK then
       BEGIN
       MessageAlerte('Suppression impossible') ; exit;
       END;
       ExecuteSQL('DELETE FROM F3310_VALRUB WHERE F10_MILLESIME="'+LiasseEnv.MILLESIME+'"'
             + ' AND F10_NUMAUTO=1 AND F10_DATEEFFET="'+ EvtSup.Datedebut+'"');
      EvtSup.Free;
      L.GotoLeBookmark(i);
    END;
    L.ClearSelected;
    LiasseEnv.DateEncours := Q.FindField('CA3_DATEDEBUT').AsDateTime;
  END;
  FiniMove;
  Q1 := Opensql ('SELECT CA3_DATEDEBUT FROM CA3_DECLARATION ORDER BY CA3_DATEDEBUT DESC', TRUE);
  if not Q1.EOF then
  LiasseEnv.DateEncours := Q1.FindField ('CA3_DATEDEBUT').asdatetime;
  Ferme (Q1);

  Date :=  USDATETIME(LiasseEnv.DateEncours);
  Date1 := ReadTokenPipe(Date,slach);
  Date1 := Date1+ ReadTokenPipe(Date,slach);
  Date1 := Date1+ copy (Date, 2, 2);
  ExecuteSQL ('Update CA3_PARAMETRE set CP3_DATEENCOURS ="' + USDATETIME(LiasseEnv.DateEncours)+'",'
           + 'CP3_FPERIODE ="' + Date1 + '"');
end;

procedure AGLSupprimeListCa3(parms: array of variant; nb: integer );
var
  F : TForm;
  Liste : THDBGrid;
  Query : TQuery;
begin
  F:=TForm(Longint(Parms[0])) ;
  if F=Nil then exit ;
  Liste:=THDBGrid(F.FindComponent('FListe') );
  if Liste=Nil then exit;
  Query:=TQuery(F.FindComponent('Q')) ;
  if (Query=Nil) then exit;
  SupprimeListeEnregca3(Liste,Query, parms[1]);
end;

// fermeture

type
TFouv = class
    Exercice  : string;
    LaTable   : string;
    Datedebut : string;
    private
    Fermer: Boolean ;
    Procedure OuvrirOuFermer;
    procedure ModifieCa3(L : THDBGrid; Q :TQuery) ;
end;
var  Evtouv : TFouv;

Procedure TFouv.OuvrirOuFermer;
BEGIN
if Fermer then
   BEGIN
      if ExecuteSql('UPDATE '+LaTable+' SET CA3_STATUT="-", CA3_DATEFERMETURE="'+UsDateTime(iDate1900)+'" '+
                    ' Where CA3_EXERCICE="'+Exercice+'" and CA3_Datedebut="' + Datedebut +'"')<>1 then V_PGI.IoError:=oeUnknown ;
   END else
   BEGIN
      if ExecuteSql('UPDATE '+LaTable+' SET CA3_STATUT="X", CA3_DATEFERMETURE="'+UsDateTime(Date)+'" '+
                    'Where CA3_EXERCICE="'+Exercice+'" and CA3_Datedebut="' + Datedebut +'"')<>1 then V_PGI.IoError:=oeUnknown ;
   END ;
END ;


Procedure TFouv.ModifieCa3(L : THDBGrid; Q :TQuery) ;
Var i                        : Byte ;
    Msg                      : string;
    valRub                   : string;
BEGIN

  if (L.NbSelected=0) and (not L.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;

if Fermer then Msg:='1;?caption?;Désirez vous ré-ouvrir les déclarations sélectionnées?;Q;YNC;N;C;'
else Msg:='2;?caption?;Désirez vous fermer la déclaration sélectionnée?;Q;YNC;N;C;';
if HShowMessage(Msg,'','')<>mrYes then exit ;

if L.AllSelected then
BEGIN
    Q.First;
    while Not Q.EOF do
    BEGIN
      MoveCur(False);
        Exercice  :=Q.FindField('CA3_EXERCICE').AsString;
        Datedebut :=USDATETIME(Q.FindField('CA3_DATEDEBUT').AsDateTime);
        if Transactions(OuvrirOuFermer,5)<>oeOK then
        BEGIN
          MessageAlerte('Ouverture/Fermeture impossible' ) ;
          Break ;
        END;
      Q.Next;
    END;
    L.AllSelected:=False;
END
ELSE
BEGIN
    for i:=0 to L.NbSelected-1 do
      BEGIN
        L.GotoLeBookMark(i) ;
        Exercice  :=Q.FindField('CA3_EXERCICE').AsString;
        Datedebut :=USDATETIME(Q.FindField('CA3_DATEDEBUT').AsDateTime);
        if Transactions(OuvrirOuFermer,5)<>oeOK then
        BEGIN
          MessageAlerte('Ouverture/Fermeture impossible' ) ;
          Break ;
        END
      END ;
END;
L.ClearSelected ;
     if not Fermer then
     begin
          valRub := fct_LEC_RUBDICO('05035');
          valRub := ReadTokenPipe (valRub, ',');

          if ExecuteSql('UPDATE '+LaTable+' SET CA3_TVAPAYER=' + valRub +
                    'Where CA3_EXERCICE="'+Exercice+'" and CA3_Datedebut="' + USDATETIME(LiasseEnv.DATEENCOURS)+'"')<>1 then V_PGI.IoError:=oeUnknown ;
     end;
END ;


procedure AGLOuvrFermListCa3(parms: array of variant; nb: integer );
var
  F : TForm;
  Liste   : THDBGrid;
  Query   : TQuery;
  fer  : string;

begin
  F:=TForm(Longint(Parms[0])) ;
  if F=Nil then exit ;
  Liste:=THDBGrid(F.FindComponent('FListe') );
  if Liste=Nil then exit;
  Query:=TQuery(F.FindComponent('Q')) ;
  if (Query=Nil) then exit;
  Evtouv := TFouv.Create;
  fer := parms[1];
  if fer = 'FALSE' then
  Evtouv.Fermer := FALSE
  else
  Evtouv.Fermer := TRUE;

  Evtouv.LaTable:='CA3_DECLARATION';
  EvtOuv.ModifieCa3(Liste,Query);
  EvtOuv.Free;
end;

procedure AGLCreatvalrub(parms: array of variant; nb: integer );
var
sSQL      : string;
Date1     : string;
F         : TForm;
begin
     F:=TForm(Longint(Parms[0])) ;
     if F=Nil then exit ;
     Date1 := parms[1];

     sSQL := 'SELECT * FROM '+LiasseEnv.FICHIERVALRUB;
     sSQL := sSQL + ' WHERE '+LiasseEnv.FVAL+'_MILLESIME = '+'"'+LiasseEnv.MILLESIME+'"';
     sSQL := sSQL + ' AND '+LiasseEnv.FVAL+'_NUMAUTO =1 ';
     sSQL := sSQL + ' AND '+LiasseEnv.FVAL+'_DATEEFFET="'+USDATETIME(StrToDate(Date1))+'"';

     if not ExisteSQL(sSQL) then
     begin
          vb_FermeDictionnaire;
          fct_OUV_RUBDICO(VLIASSE^.RepDeclaration+ '\'+LiasseEnv.MILLESIME);
               //FONCTION SPECIFIQUE DU MOTEUR
          vb_SetUserInfoBd (InfoBdPgi);

          //MAJ DU CODEPER DANS LE DICO
          fct_MAJ_CODEPER_DICO(V_PGI.CodeSociete);

          fct_CREA_VALRUB;
     end;
end;

procedure CA3PrintDeclaTVA (Vr : Boolean);
var sSQL,sSQLSel      : string;
sRetour               :PChar;
begin

  sSQL := ' WHERE '+LiasseEnv.FVAL+'_MILLESIME = "'+LiasseEnv.MILLESIME+'"'+
          ' AND '+ LiasseEnv.FVAL+'_DATEEFFET = "'+USDATETIME(LiasseEnv.DATEENCOURS)+'"';
  sSQLSel := ' AND '+LiasseEnv.FVAL+'_MILLESIME = "'+LiasseEnv.MILLESIME+'"'+
          ' AND '+ LiasseEnv.FVAL+'_DATEEFFET = "'+USDATETIME(LiasseEnv.DATEENCOURS)+'"';

//  LanceEdition(2,'F0000_EDITION',sSQLSel,sSQL,LiasseEnv.NOMDECLARATION,'',FALSE);
LanceEdition(2,'F0000_EDITION','EDITIONS CA3', VLIASSE^.RepDeclaration ,'',Vr);

end;

procedure CA3PrintDecla (parms: array of variant; nb: integer);
var sSQL,sSQLSel      : string;
Date1                 : string;
F                     : TForm;
sRetour               : PChar;
SaveDateEncours       : TDateTime;
begin
  F:=TForm(Longint(Parms[0])) ;
  if F=Nil then exit ;
  sRetour := nil;
  vb_GetRubriqueValue(100,@sRetour,0,LiasseEnv.cDeviseExercice);
  Date1 := sRetour;

SaveDateEncours := LiasseEnv.DateEncours;
LiasseEnv.DateEncours := StrToDate(Date1);
LanceEdition(2,'F0000_EDITION','EDITIONS CA3', VLIASSE^.RepDeclaration ,'o',FALSE);
LiasseEnv.DateEncours :=  SaveDateEncours;
end;

FUNCTION TOF_INFOCA3.QUELEXODTCA3(DD : TDateTime) : String ;
Var i : Integer ;
begin
Result:='';
If (dd>=VH^.EnCours.Deb) and (dd<=VH^.EnCours.Fin) then Result:=VH^.EnCours.Code else
   If (dd>=VH^.Suivant.Deb) and (dd<=VH^.Suivant.Fin) then Result:=VH^.Suivant.Code Else
      If (dd>=VH^.Precedent.Deb) and (dd<=VH^.Precedent.Fin) then Result:=VH^.Precedent.Code Else
      BEGIN
         For i:=1 To 5 Do
           BEGIN
           If (dd>=VH^.ExoClo[i].Deb) And (dd<=VH^.ExoClo[i].Fin)then BEGIN Result:=VH^.ExoClo[i].Code ; Exit ; END ;
           END ;
      END ;
end ;


Initialization
registerclasses([TOF_INFOCA3]);
RegisterAglProc( 'SupprimeListeCa3', TRUE , 1, AGLSupprimeListCa3);
RegisterAglProc( 'OuvrFermListCa3', TRUE , 1, AGLOuvrFermListCa3);
RegisterAglProc( 'OuvrFermListCa3', TRUE , 1, AGLOuvrFermListCa3);
RegisterAglProc( 'Creatvalrub', TRUE , 1, AGLCreatvalrub);
RegisterAglProc( 'PrintDecla', TRUE , 1, CA3PrintDecla);
end.
