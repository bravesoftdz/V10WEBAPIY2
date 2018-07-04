{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/01/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AUGMEXCLUS (AUGMEXCLUS)
Mots clefs ... : TOM;AUGMEXCLUS
*****************************************************************}
Unit UTomAugmParam ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList,
     FE_Main,
{$else}
     eFiche,
     eFichList,
     MainEAGL, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM,
     Hqry,
     HTB97,
     PGOutils2,
     HSysMenu,
     UTob ;

Type
  TOM_AUGMPARAM = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    GCritere : THGrid;
    procedure VoirDetail(Sender : TObject);
    procedure RemplirGrille;
    procedure ChangeOnglet(Sender : TObject);
    procedure ModifSupCritere(Sender : TObject);
    procedure AjoutCritere(Sender : TObject);
    end ;

Implementation

procedure TOM_AUGMPARAM.OnNewRecord ;
var IMax : Integer;
    QQuery : TQuery;
    Code : String;
begin
  Inherited ;
            SetControlEnabled('BDETAIL',False);
            SetField('PAP_ETATINTAUGM','012');
            SetField('PAP_ACTIF','X');
            QQuery := OpenSQL('SELECT MAX(PAP_CRITEREAUGM) FROM AUGMPARAM', TRUE);
            if not QQuery.Eof then
            begin
                  if QQuery.Fields[0].AsString <> '' then Imax := StrToInt(QQuery.Fields[0].AsString)
                  else Imax := 0;
                  IMax := IMax + 1;
                  Code := ColleZeroDevant(IMax, 5);
                  SetField('PAP_CRITEREAUGM', Code);
            end;
            Ferme(QQuery);
end ;

procedure TOM_AUGMPARAM.OnDeleteRecord ;
begin
  Inherited ;
  ExecuteSQL('DELETE FROM AUGMEXCLUS WHERE PAE_CRITEREAUGM="'+GetField('PAP_CRITEREAUGM')+'"');
end ;

procedure TOM_AUGMPARAM.OnUpdateRecord ;
begin
  Inherited ;
     SetControlEnabled('BDETAIL',True);
     If TPageControl(GetControl('PAGES')).ActivePageIndex = 1 then SetControlEnabled('BAJOUT',True);
end ;

procedure TOM_AUGMPARAM.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMPARAM.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMPARAM.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMPARAM.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_AUGMPARAM.OnArgument ( S: String ) ;
var BDetail,BAjout : TToolBarButton97;
Onglet : TTabSheet;
begin
  Inherited ;
  BDetail := TToolBarButton97(GetControl('BDETAIL'));
  If BDetail <> Nil then BDetail.OnClick := VoirDetail;
  GCritere := THGrid(GetControl('GCRITERES'));
  GCritere.ColWidths[0] := -1;
  GCritere.ColWidths[1] := -1;
  GCritere.ColWidths[2] := 200;
  GCritere.ColWidths[3] := 70;
  GCritere.ColAligns[3] := TaCenter;
  GCritere.ColWidths[4] := 200;
  GCritere.ColFormats[2] := 'CB=PGAUGMCRITERES';
  Onglet := TTabSheet(GetControl('PCRITERES'));
  If Onglet <> Nil then Onglet.OnShow := ChangeOnglet;
  Onglet := TTabSheet(GetControl('PGeneral'));
  If Onglet <> Nil then Onglet.OnShow := ChangeOnglet;
  GCritere.OnDblClick := ModifSupCritere;
  BAjout := TToolBarButton97(GetControl('BAJOUT'));
  If BAjout <> nil then BAjout.OnClick := AjoutCritere;
  SetControlEnabled('BAJOUT',False);
end ;

procedure TOM_AUGMPARAM.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_AUGMPARAM.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMPARAM.VoirDetail(Sender : TObject);
begin
     AglLanceFiche('PAY', 'MUL_AUGMEXCLUS', '', '', GetField('PAP_CRITEREAUGM'));
end;

procedure TOM_AUGMPARAM.AjoutCritere(Sender : TObject);
begin
     AglLanceFiche('PAY', 'AUGMEXCLUS', '', '', GetField('PAP_CRITEREAUGM')+';ACTION=CREATION');
     RemplirGrille;
end;

procedure TOM_AUGMPARAM.ModifSupCritere(Sender : TObject);
var NumOrdre : String;
begin
     If DS.State = DSInsert then Exit;
     NumOrdre := GCritere.CellValues[0,GCritere.Row];
     If NumOrdre = '' then exit;
     AglLanceFiche('PAY', 'AUGMEXCLUS', '',NumOrdre+';'+GetField('PAP_CRITEREAUGM') , '');
     RemplirGrille;
end;

procedure TOM_AUGMPARAM.ChangeOnglet (Sender : TObject);
begin
        If TTabSheet(Sender).Name = 'PCRITERES' then
        begin
             if DS.State = dsInsert then
             begin
                  SetControlEnabled('BAJOUT',False);
//                  TPageControl(GetControl('PAGES')).ActivePage := TTabSheet(GetControl('PGeneral'));
                  PGIBox('Vous devez valider le nouveau paramètre d''exclusion avant de saisir les critères',Ecran.Caption);
                  Exit;
             end;
             RemplirGrille;
             SetControlEnabled('BAJOUT',True);
        end
        else SetControlEnabled('BAJOUT',False);
end;

procedure TOM_AUGMPARAM.RemplirGrille;
var TobCritere : Tob;
    StChampGrid : String;
    i : Integer;
    TypeValeur : String;
    Q : TQuery;
begin
     Q := OpenSQL('SELECT PAE_ORDRE,PAE_LIBELLE,PAE_CRITEREAUGEX,PAE_OPERATTEST,PAE_VALEUREX'+
     ' FROM AUGMEXCLUS WHERE PAE_CRITEREAUGM="'+GetField('PAP_CRITEREAUGM')+'"',True);
     TobCritere := Tob.create('Lescriteres',Nil,-1);
     TobCritere.LoadDetailDB('LesCriteres','','',Q,False);
     Ferme(Q);
     For i := 0 to TobCritere.Detail.Count - 1 do
     begin
          Q := OpenSQL('SELECT CO_ABREGE FROM COMMUN WHERE CO_CODE="'+TobCritere.Detail[i].GetValue('PAE_CRITEREAUGEX')+'" AND CO_TYPE="PGM"',True);
          If Not Q.Eof then TypeValeur := Q.FindField('CO_ABREGE').AsString
          else TypeValeur := 'DATA';
          Ferme(Q);
          If TypeValeur <> 'DATA' then TobCritere.Detail[i].PutValue('PAE_VALEUREX',RechDom(TypeValeur,TobCritere.Detail[i].GetValue('PAE_VALEUREX'),False));
     end;
     GCritere.RowCount := TobCritere.Detail.Count + 1;
     If GCRitere.RowCount = 1 then GCRitere.RowCount := 2;
     StChampGrid := 'PAE_ORDRE;PAE_LIBELLE;PAE_CRITEREAUGEX;PAE_OPERATTEST;PAE_VALEUREX';
     TobCritere.PutGridDetail(GCritere,False,False,StChampGrid,False);
     TobCritere.Free;
     GCRitere.FixedRows := 1; 
end;

Initialization
  registerclasses ( [ TOM_AUGMPARAM ] ) ;
end.

