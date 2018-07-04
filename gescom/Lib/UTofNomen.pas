unit UTofNomen;

interface
uses  M3FP, StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,
{$IFDEF EAGLCLIENT}
      MaineAGL,
      eFichList,
{$ELSE}
      Fe_Main, db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}MajTable,
			FichList,
{$ENDIF}
{$IFDEF BTP}
      UTOFBTANALDEV,
{$ENDIF}
      UTOF, AglInit, NomenLig,
      NomenAPlat, CasEmplois;

Type
     TOF_NomenEnt = Class (TOF)
       private

       public
       procedure OnArgument (params : string) ; override ;
       procedure OnCancel ; override ;
       procedure OnNew ; override ;
       procedure OnDelete ; override ;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
     END ;

     TOF_NomenLig = Class (TOF)
       private

       public
       procedure OnCancel ; override ;
       procedure OnNew ; override ;
       procedure OnDelete ; override ;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
     END ;

implementation
uses BTPUtil,HrichOle;
/////////////////////////////////////////////////////////////////////////////
procedure TOF_NomenEnt.OnCancel;
begin
    Inherited;
end;

procedure TOF_NomenEnt.OnNew;
begin
    Inherited;
end;

procedure TOF_NomenEnt.OnDelete;
begin
    Inherited;
end;

procedure TOF_NomenEnt.OnLoad;
begin
    Inherited;
end;

procedure TOF_NomenEnt.OnUpdate;
begin
    Inherited;
if (LaTOB<>Nil) then TheTob := LaTOB ;
end;

/////////////////////////////////////////////////////////////////////////////
procedure TOF_NomenLig.OnCancel;
begin
    Inherited;
end;

procedure TOF_NomenLig.OnNew;
begin
    Inherited;

end;

procedure TOF_NomenLig.OnDelete;
begin
    Inherited;
end;

procedure TOF_NomenLig.OnLoad;
begin
    Inherited;
end;

procedure TOF_NomenLig.OnUpdate;
begin
    Inherited;
if (LaTOB<>Nil) then TheTob := LaTOB ;
end;

/////////////////////////////////////////////////////////////////////////////
procedure TOF_NomenEnt_POPZ( Parms: array of variant; nb: integer ) ;
var
    F : TForm;
    Option : Integer;
    Valeur : string;
    st1, st2 : string;
    // Modif BTP
{$IFDEF BTP}
    st3 : String;
{$ENDIF}
begin
F := TForm(Longint(Parms[0]));
Option := Integer(Parms[1]);
Valeur := String(Parms[2]);
case Option of
1 : begin
{$IFDEF CHR}
    st1:= Valeur;
    st2:= Valeur;
{$ELSE}
    st1 := THEdit(F.FindComponent('GNE_ARTICLE')).Text;
    st2 := THEdit(F.FindComponent('GNE_NOMENCLATURE')).Text;
{$ENDIF}
{$IFDEF BTP}
    St3 := THEdit(F.FindComponent('GNE_QTEDUDETAIL')).Text;
    Entree_NomenLig(['CASEMPLOIS=OUI', st1, st2, st3,TFFicheListe(F).fTypeAction], 4);
//    THEdit(F.FindComponent('GNE_DATEMODIF')).Text := DateToStr (now);
{$ELSE}
    Entree_NomenLig(['CASEMPLOIS=OUI', st1, st2], 3) ;
{$ENDIF}
    end;
2 : begin
{$IFDEF BTP}
       // Analyse BTP
        st1 := THEdit(F.FindComponent('GNE_ARTICLE')).Text;
        st2 := THEdit(F.FindComponent('GNE_NOMENCLATURE')).Text;
        st3 := THEdit(F.FindComponent('GNE_LIBELLE')).Text;
        EntreeAnalyseOuvrageBib (st1,St2,st3);
{$ELSE}
        FNomenAPlat := TFNomenAPlat.Create(Application);
        Try
        FNomenAPlat.GNE_ARTICLE.Text := THEdit(F.FindComponent('GNE_ARTICLE')).Text;
        // FNomenAPlat.GNE_TYPNOMENC.Text := THDBValComboBox(F.FindComponent('GNE_TYPNOMENC')).Text;
        FNomenAPlat.GNE_NOMENCLATURE.Text := THEdit(F.FindComponent('GNE_NOMENCLATURE')).Text;
        FNomenAPlat.GNE_LIBELLE.Text := THEdit(F.FindComponent('GNE_LIBELLE')).Text;
        FNomenAPlat.ShowModal;
        finally
        FNomenAPlat.free;
        end;
{$ENDIF}
    end;
3 : begin
    st1 := THEdit(F.FindComponent('GNE_NOMENCLATURE')).Text;
    Entree_CasEmploi(['NOMENCLATURE=NON', st1], 2)
    end;
end;
end;

/////////////////////////////////////////////////////////////////////////////
procedure InitTOFNomen();
begin
 RegisterAglProc( 'NomenEnt_POPZ', True , 1, TOF_NomenEnt_POPZ);
end;

procedure TOF_NomenEnt.OnArgument(params: string);
begin
  inherited;
  //
  AppliqueFontDefaut (THRichEditOle(GetControl('GNE_BLOCNOTE')));
end;

Initialization
registerclasses([TOF_NomenEnt, TOF_NomenLig]) ;
InitTOFNomen();

end.
