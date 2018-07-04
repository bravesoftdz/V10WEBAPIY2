unit UTomCodeCpta;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      DB,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main,
{$ENDIF}
      UTOM,Utob;


Type
     TOM_CodeCtpa = Class (TOM)
     private
     public
       procedure OnUpdateRecord ; override ;
     END ;

const
// libellés des messages
TexteMessage: array[1..15] of string 	= (
          {1}  'Cette fiche existe déjà, vous devez la modifier'
          {2} ,''
          {3} ,''
          {4} ,''
          {5} ,''
          {6} ,''
          {7} ,''
          {8} ,''
          {9} ,''
         {10} ,''
         {11} ,''
         {12} ,''
         {13} ,''
         {14} ,''
         {15} ,''
              );

implementation

procedure TOM_CodeCtpa.OnUpdateRecord;
var QQ      :TQuery;
    RangMax :String;
    Rang    :Integer;
begin
Inherited;
//Incrémenter le champ Numerocom
if (DS.State in [dsInsert]) then
    BEGIN
    QQ := OpenSQL ('SELECT MAX(GCP_RANG) FROM CODECPTA',True);
    if Not QQ.EOF then
        BEGIN
        RangMax := QQ.Fields[0].AsString;
        if RangMax = '' then Rang := 1 else Rang := StrToInt(RangMax)+1;
        Setfield('GCP_RANG', Rang);
        Ferme(QQ);
        END else
        BEGIN
        Ferme(QQ);
        SetFocusControl('GCP_COMPTATIERS');
        LastError := 1;
        LastErrorMsg := TexteMessage[LastError];
        Exit;
        END;
    END;
END;

end.
 
