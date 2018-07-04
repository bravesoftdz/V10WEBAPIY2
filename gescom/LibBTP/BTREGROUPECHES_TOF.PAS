{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/12/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREGROUPECHES ()
Mots clefs ... : TOF;BTREGROUPECHES
*****************************************************************}
Unit BTREGROUPECHES_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     HTB97,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ;

Type
  TOF_BTREGROUPECHES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fTerme,AffAire : string;
    procedure BOuvrirClick (Sender : TObject);
  end ;

Implementation
uses AffEcheanceUtil;


procedure TOF_BTREGROUPECHES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREGROUPECHES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREGROUPECHES.OnUpdate ;
var QQ : TQuery;
begin
  Inherited ;
  AffAire := GetControlText('AFA_AFFAIRE');
  QQ := OpenSQL('SELECT AFF_TERMEECHEANCE FROM AFFAIRE WHERE AFF_AFFAIRE="'+AffAire+'"',True,1,'',true);
  if not QQ.eof then fTerme := QQ.fields[0].AsString;
  ferme (QQ);
end ;

procedure TOF_BTREGROUPECHES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREGROUPECHES.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolbarButton97(getControl('BOuvrir')).OnClick  := BOuvrirClick;
end ;

procedure TOF_BTREGROUPECHES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTREGROUPECHES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREGROUPECHES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTREGROUPECHES.BOuvrirClick(Sender: TObject);
var F : TForm;
    Q : TQuery;
    TOBEches,TB : TOB;
    i,iFirst,NIndice : Integer;
    first : boolean;
    preums,last,courant : Integer;
    MontantGlob,MontantGlobDev : Double;
    DateFin : TDateTime;
    contigous,OKOK : boolean;

begin
  F := Ecran;

  if(TFMul(F).FListe.NbSelected=0)and(not TFMul(F).FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné','');
    exit;
  end;

  if PGIAsk('ATTENTION : Vous allez regrouper les échéances sélectionnées.#13#10 Confirmez-vous ?')<>Mryes then Exit;

  first := True;
  iFirst := -1;
  MontantGlob := 0;
  courant := 0;
  contigous := True;
  OKOK := false;

  TOBEches := TOB.create ('LES ECHES',nil,-1);
  Q := OpenSQL('SELECT *,"-" AS IMPACTE FROM FACTAFF WHERE AFA_AFFAIRE="'+AffAire+'" AND AFA_ECHEFACT="-"',True,-1,'',true);
  if not Q.eof then
  begin
    TOBEches.LoadDetailDB ('FACTAFF','','',Q,false);
  end;
  Ferme(Q);

  TRY

    if TFMul(F).Fliste.AllSelected then
    BEGIN
      Q:=TFMul(F).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        TB := TOBeches.FindFirst(['AFA_NUMECHE'],[Q.findField('AFA_NUMECHE').AsInteger],true);
        if TB <> nil then TB.SetString('IMPACTE','X');
        if first then
        begin
          preums := Q.findField('AFA_NUMECHE').AsInteger;
          if TB <> nil then iFirst := TB.GetIndex;
          first := false;
        end;

        MontantGlobDev := MontantGlobDev +Q.findField('AFA_MONTANTECHEDEV').AsFloat;
        MontantGlob := MontantGlob +Q.findField('AFA_MONTANTECHE').AsFloat;
        last := Q.findField('AFA_NUMECHE').AsInteger;
        DateFin := TFMul(F).Fliste.datasource.dataset.findField('AFA_DATEFINFAC').AsDateTime;
        Q.NEXT;
      END;
      TFMul(F).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
        TFMul(F).Fliste.GotoLeBookmark(i);
        if courant +1 <> TFMul(F).Fliste.datasource.dataset.findField('AFA_NUMECHE').AsInteger then
        begin
          contigous := false;
        end else
        begin
          courant := TFMul(F).Fliste.datasource.dataset.findField('AFA_NUMECHE').AsInteger;
        end;
        TB := TOBeches.FindFirst(['AFA_NUMECHE'],[TFMul(F).Fliste.datasource.dataset.findField('AFA_NUMECHE').AsInteger],true);
        if TB <> nil then TB.SetString('IMPACTE','X');
        if first then
        begin
          if TB <> nil then iFirst := TB.GetIndex;
          preums := TFMul(F).Fliste.datasource.dataset.findField('AFA_NUMECHE').AsInteger;
          first := false;
        end;
        if not contigous then break;

        MontantGlobDev := MontantGlobDev +TFMul(F).Fliste.datasource.dataset.findField('AFA_MONTANTECHEDEV').AsFloat;
        MontantGlob := MontantGlob +TFMul(F).Fliste.datasource.dataset.findField('AFA_MONTANTECHE').AsFloat;
        last := TFMul(F).Fliste.datasource.dataset.findField('AFA_NUMECHE').AsInteger;
        DateFin := TFMul(F).Fliste.datasource.dataset.findField('AFA_DATEFINFAC').AsDateTime;
      end;
      TFMul(F).Fliste.ClearSelected;
    END;
    
    if not contigous then
    begin
      PGIError('ERREUR : Les sélections doivent être contigues');
      Exit;
    end;
    if preums = last then exit;
    OKOK := true;

    if iFirst <> -1 then
    begin
      for I := iFirst to TOBEches.Detail.count -1 do
      begin
        TB := TOBEches.detail[I];
        if I = iFirst then
        begin
          TB.PutValue('AFA_DATEFINFAC',DateFin);
          TB.SetDouble('AFA_MONTANTECHEDEV',MontantGlobDev);
          TB.SetDouble('AFA_MONTANTECHE',MontantGlob);
          TB.SetString('AFA_LIBELLEECHE',ConstitueLibEcheContrat(TB));
          if fTerme = 'ECH' then TB.PutValue('AFA_DATEECHE',DateFin);
          NIndice := TB.GetInteger('AFA_NUMECHE');
          TB.UpdateDB;
        end else
        begin
          if TB.GetString('IMPACTE')<>'X' then
          begin
            Inc(NIndice);
            TB.SetInteger('AFA_NUMECHE',Nindice);
            TB.SetInteger('AFA_NUMECHEBIS',Nindice);
            TB.SetString('AFA_LIBELLEECHE',ConstitueLibEcheContrat(TB));
            TB.UpdateDB;
          end else
          begin
            TB.DeleteDB;
          end;
        end;
      end;
    end;
    
  FINALLY
    TOBEches.Free;
    if OKOK then Ecran.Close;
  END;
end;

Initialization
  registerclasses ( [ TOF_BTREGROUPECHES ] ) ;
end.
