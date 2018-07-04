{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SALAIREMOYEN () Salaires moyens calculés 
Mots clefs ... : TOF;SALAIREMOYEN
*****************************************************************
PT1   14/05/2007 :  NA  Suppression des calculs sélectionnés}

Unit SALAIREMOYEN_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,

{$IFNDEF EAGLCLIENT}
     db, 
     dbtables,
     Mul,
{$ELSE}
     uTob,
     eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     P5Def,
     HTB97,
     HmsgBox,
     HCtrls,
     UTOF ;

Type
  TOF_SALAIREMOYEN = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure Suppressionsalmoyen(sender: TObject);      // pt1
  end ;

Implementation

procedure TOF_SALAIREMOYEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SALAIREMOYEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SALAIREMOYEN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SALAIREMOYEN.OnLoad ;
begin
   Inherited;
 end ;

procedure TOF_SALAIREMOYEN.OnArgument (S : String ) ;
VAR
Anneeprec, num : integer;
AnneeA, MoisM, jour : word;
jourj, zdate,datecal : Tdatetime;
btn : TToolBarButton97;  // pt1
begin
  Inherited ;
  Jourj := Date;
  decodedate(Jourj,AnneeA, MoisM, Jour);
  Anneeprec := AnneeA - 1;
  zdate := Encodedate(anneeprec, moisM, Jour);
  datecal := finannee(Zdate);
  SetControlText ('PSY_DATECALC', DateTimeToStr(datecal));

  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSY_TRAVAILN' + IntToStr(Num)), GetControl('TPSY_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSY_CODESTAT'), GetControl('TPSY_CODESTAT'));

  // pt1
  Btn := TToolBarButton97(GetControl('B_SUPPRIMER'));
  if Btn <> nil then Btn.OnClick := Suppressionsalmoyen;
  // fin pt1
end ;

procedure TOF_SALAIREMOYEN.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_SALAIREMOYEN.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_SALAIREMOYEN.OnCancel () ;
begin
  Inherited ;
end ;

// pt1
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 14/05/2007
Modifié le ... :   /  /    
Description .. : Suppression des salaires moyens pour les salariés  
Suite ........ : sélectionnés
Mots clefs ... : 
*****************************************************************}
procedure TOF_SALAIREMOYEN.Suppressionsalmoyen(sender : TObject) ;
var
toutselect : boolean;
tot, i : integer;
salarie , codecalc, st: string;
datecalc : Tdatetime;

begin
if (TFMul(Ecran).FListe.nbSelected > 0) or (TFMul(Ecran).FListe.AllSelected) then
    if PgiAsk('Etês-vous sûr de vouloir supprimer les salaires moyens sélectionnés?', Ecran.caption) = mrYes then
    begin
      ToutSelect := False;
      If TFMul(Ecran).FListe.AllSelected then
      Begin
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Fetchlestous;
        {$ENDIF}
        ToutSelect := True;
        Tot := TFMul(Ecran).Q.RecordCount;

        TFMul(Ecran).Q.First;
      End
      else Tot := TFMul(Ecran).FListe.NbSelected;

      { Balayage des lignes sélectionnées }
      for i := 0 to Tot - 1 do
      begin
        {$IFDEF EAGLCLIENT}
        if not ToutSelect then  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
        {$ENDIF}
        if TFMul(Ecran).FListe.NbSelected > 0 then TFMul(Ecran).FListe.GotoLeBookmark(i);
        salarie    := TFmul(Ecran).Q.FindField('PSY_SALARIE').asstring;
        Datecalc   := TFmul(Ecran).Q.FindField('PSY_DATECALC').AsDateTime;
        Codecalc   := TFmul(Ecran).Q.FindField('PSY_CODECALSAL').Asstring;

        St := 'DELETE FROM SALAIREMOYEN '+
              ' WHERE PSY_SALARIE = "'+salarie+'" ' +
              'AND PSY_DATECALC ="'+USDateTime(Datecalc)+'" '+
              'AND PSY_CODECALSAL ="'+Codecalc+'" ';

        ExecuteSql(St);
        if ToutSelect then TFMul(Ecran).Q.Next;
      end;
    TFMul(Ecran).FListe.ClearSelected;
    TFMul(Ecran).BCherche.Click;
    End;
  
end;
// fin pt1

Initialization
  registerclasses ( [ TOF_SALAIREMOYEN ] ) ; 
end.
