{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/02/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTARTPARC_MUL ()
Mots clefs ... : TOF;BTARTPARC_MUL
*****************************************************************}
Unit BTARTPARC_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
     HDB,EdtREtat,
{$else}
     MaineAGL,
     eMul,MaineAGL,
{$ENDIF}
     uTob,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HTB97, 
     HMsgBox,
     AglInit, 
     UTOF ;

Type
  TOF_BTARTPARC_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    public
    	MultiSelect : boolean;
    	procedure RemplacecolonnesFamilles;
      procedure BOuvrirClick (Sender : Tobject);
  end ;

Implementation

procedure TOF_BTARTPARC_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTARTPARC_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTARTPARC_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTARTPARC_MUL.OnLoad ;
begin
  Inherited ;
  RemplacecolonnesFamilles;
end ;

procedure TOF_BTARTPARC_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  MultiSelect := (Pos('MULTISELECTION',S) > 0);
  {$IFNDEF EAGLCLIENT}
  TFMul(Ecran).FListe.MultiSelection := MultiSelect;
 {$ELSE}
  TFMul(Ecran).Fliste.MultiSelect := MultiSelect;
 {$ENDIF}
  if MultiSelect then TToolBarButton97(GetControl('BOUVRIR')).OnClick := BOuvrirClick;
end ;

procedure TOF_BTARTPARC_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTARTPARC_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTARTPARC_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTARTPARC_MUL.RemplacecolonnesFamilles;
var i : integer;
		Gr : THDbgrid;
    stChamp,Libelle : string;
begin
	Gr := TFMul(ecran).fliste;
	For i:=0 to Gr.Columns.Count-1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.Columns[i].FieldName);
    if copy(UpperCase (stChamp),1,7)='LIBPARC' then
    begin
      libelle := RechDom('BTLIBARTPARC','PA'+Copy(stChamp,8,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end;
  end;
end;

procedure TOF_BTARTPARC_MUL.BOuvrirClick(Sender: Tobject);
var iInd : integer;
    FieldArticle : string;
    TobRetour,TobRF : TOB;
begin
  FieldArticle := 'GA_ARTICLE';
  with TFMul(Ecran) do
  begin
    if not MultiSelect then
    begin
      Retour := Q.FindField(FieldArticle).AsString;
    end else
    begin
      TobRetour := TOB.Create('',nil,-1);
      if not FListe.AllSelected then
      begin
        if FListe.NbSelected = 0 then
        begin
          TobRF := TOB.Create('',TobRetour,-1);
          TobRF.AddChampSupValeur('ARTICLE',Q.FindField(FieldArticle).AsString);
        end else
        begin
          for iInd := 0 to FListe.NbSelected -1 do
          begin
            FListe.GotoLeBookMark(iInd);
    {$IFDEF EAGLCLIENT}
            Q.TQ.Seek (FListe.Row-1) ;
    {$ENDIF}
            TobRF := TOB.Create('',TobRetour,-1);
            TobRF.AddChampSupValeur('ARTICLE',Q.FindField(FieldArticle).AsString);
          end;
        end;
      end else
      begin
        Q.First;
        While not Q.Eof do
        begin
          TobRF := TOB.Create('',TobRetour,-1);
          TobRF.AddChampSupValeur('ARTICLE',Q.FindField(FieldArticle).AsString);
          Q.Next;
        end;
      end;
      TheTob := TobRetour;
      if TheTob.Detail.Count = 1 then Retour := TheTob.Detail[0].GetValue('ARTICLE');
    end;
    Close;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTARTPARC_MUL ] ) ; 
end.
