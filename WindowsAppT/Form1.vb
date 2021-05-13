Imports Octave.NET
Public Class Form1
    Dim vectory(10) As Double
    Dim vectort(10) As Double
    Dim vectory2(10) As Double
    Dim vectort2(10) As Double
    Dim vectory3(10) As Double
    Dim vectort3(10) As Double
    Dim vectory4(10) As Double
    Dim vectort4(10) As Double
    Dim i As Integer = 0
    Dim yinicial As Integer
    Dim yinicial2 As Integer
    Dim altoinicial As Integer
    Dim cantsimulaciones As Integer = 3
    Dim index_row As Integer = 0
    Dim index_row2 As Integer = 0
    Dim index_row3 As Integer = 0
    Dim index_row4 As Integer = 0
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Timer1.Interval = intervalotxt.Text
        i = 0
        Chart1.Series(0).Points.Clear()
        Chart2.Series(0).Points.Clear()
        OctaveContext.OctaveSettings.OctaveCliPath =
       "C:\Program Files\GNU Octave\Octave-6.2.0\mingw64\bin\Octave-cli.exe"
        Dim Octave = New OctaveContext()
        Octave.Execute("clear;
                        pkg load control;
                        s=tf('s');
                        m1=" & m1txt.Text & ";
                        m2=" & m2txt.Text & ";
                        b1=" & btxt.Text & ";
                        b2=" & b2txt.Text & ";
                        k1=" & k1txt.Text & ";
                        k2=" & k1txt.Text & ";
                        k3=" & k1txt.Text & ";
                        g=(m2 * s^2 + b2 * s + k3 + k2)/[(m1*s^2+k1+k3+b1 * s)*(m2 * s^2 + b2*s+k3+k2) - k3^2];

                        [y,t]=impulse(g);
                        c=length(t);
                        tiempo=t(c)*1.2;
                        [y,t]=impulse(g,tiempo,tiempo/" &
                        ctxt.Text & ");
                        
                        
                        g2=k3/[(m1*s^2+k1+k3+b1 * s)*(m2 * s^2 + b2*s+k3+k2) - k3^2];

                        g3=(m2 * s^2 + b2 * s + k3 + k2)/[(m1*s^2+k1+k3+b1)*(m2 * s^2 + b2*s+k3+k2) - k3^2];
                        g4=k3/[(m1*s^2+k1+k3+b1)*(m2 * s^2 + b2*s+k3+k2) - k3^2];


                        [y2,t2]=impulse(g2);
                        [y2,t2]=impulse(g2,tiempo,tiempo/" &
                        ctxt.Text & ");


                        [y3,t3]=impulse(g3);
                        [y3,t3]=impulse(g3,tiempo,tiempo/" &
                        ctxt.Text & ");
                            

                        [y4,t4]=impulse(g4);
                        [y4,t4]=impulse(g4,tiempo,tiempo/" &
                        ctxt.Text & ");
                            
                            
                        ")


        '[y,t]=step(g,timepo sim,rango tiempo);")
        ReDim vectory(ctxt.Text)
        ReDim vectort(ctxt.Text)
        vectory = Octave.Execute("y").AsVector
        vectort = Octave.Execute("t").AsVector

        ReDim vectory2(ctxt.Text)
        ReDim vectort2(ctxt.Text)
        vectory2 = Octave.Execute("y2").AsVector
        vectort2 = Octave.Execute("t2").AsVector

        ReDim vectory3(ctxt.Text)
        ReDim vectort3(ctxt.Text)
        vectory3 = Octave.Execute("y3").AsVector
        vectort3 = Octave.Execute("t3").AsVector

        ReDim vectory4(ctxt.Text)
        ReDim vectort4(ctxt.Text)
        vectory4 = Octave.Execute("y4").AsVector
        vectort4 = Octave.Execute("t4").AsVector

        masaPb.Width = masaReferencia.Width
        masaPb.Location = masaReferencia.Location

        masaPb2.Width = masa2Referencia.Width
        masaPb2.Location = masa2Referencia.Location

        resortePb1.Width = CajaReferencia.Width
        amortPb1.Width = CajaReferencia.Width
        resortePb2.Width = cajaReferencia2.Width
        amortPb2.Width = cajaReferencia2.Width
        Timer1.Enabled = True

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If ComboBox1.SelectedItem.ToString = "Resorte" Then

            amortPb1.BackColor = Color.Orange
            Chart1.Series(0).Points.AddXY(vectort3(i), vectory3(i))
            Chart2.Series(0).Points.AddXY(vectort4(i), vectory4(i))
            masaPb.Location = New Point(masaPb.Location.X + vectory3(i) * 100,
                    yinicial)
            masaPb2.Location = New Point(masaPb2.Location.X + vectory4(i) * 100,
                    yinicial2)

            resortePb1.Width = resortePb1.Width + vectory3(i) * 100
            amortPb1.Width = amortPb1.Width + vectory3(i) * 100


            resortePb2.Location = New Point(masaPb2.Location.X + vectory4(i),
                    resortePb2.Location.Y)
            resortePb2.Width = resortePb2.Width + vectory4(i) * 100


            resortePb3.Location = New Point(masaPb.Location.X + vectory3(i),
                    resortePb3.Location.Y)



            amortPb2.Location = New Point(masaPb2.Location.X + vectory4(i),
                    amortPb2.Location.Y)
            amortPb2.Width = amortPb2.Width + vectory4(i) * 100





        Else
            amortPb1.BackColor = Color.Fuchsia


            Chart1.Series(0).Points.AddXY(vectort(i), vectory(i))
            Chart2.Series(0).Points.AddXY(vectort2(i), vectory2(i))
            masaPb.Location = New Point(masaPb.Location.X + vectory(i) * 100,
                    yinicial)
            masaPb2.Location = New Point(masaPb2.Location.X + vectory2(i) * 100,
                    yinicial2)

            resortePb1.Width = resortePb1.Width + vectory(i) * 100
            amortPb1.Width = amortPb1.Width + vectory(i) * 100


            resortePb2.Location = New Point(masaPb2.Location.X + vectory2(i),
                    resortePb2.Location.Y)
            resortePb2.Width = resortePb2.Width + vectory2(i) * 100


            resortePb3.Location = New Point(masaPb.Location.X + vectory(i),
                    resortePb3.Location.Y)



            amortPb2.Location = New Point(masaPb2.Location.X + vectory2(i),
                    amortPb2.Location.Y)
            amortPb2.Width = amortPb2.Width + vectory2(i) * 100





        End If


        i += 1
        If i = ctxt.Text Then
            cantsimulaciones = cantsimulaciones - 1
            DataGridView1.Rows.Add()
            DataGridView1.Item(0, index_row).Value = m1txt.Text
            DataGridView2.Rows.Add()
            DataGridView2.Item(0, index_row2).Value = m2txt.Text

            'registrar imagen en la celda
            Dim captura_chart As New Bitmap(Chart1.Width, Chart1.Height)
            Chart1.DrawToBitmap(captura_chart, Chart1.DisplayRectangle)
            DataGridView1.Item(1, index_row).Value = captura_chart
            index_row += 1
            m1txt.Text = m1txt.Text * 2


            Dim captura_chart2 As New Bitmap(Chart2.Width, Chart2.Height)
            Chart2.DrawToBitmap(captura_chart2, Chart2.DisplayRectangle)
            DataGridView2.Item(1, index_row2).Value = captura_chart2
            index_row2 += 1
            m2txt.Text = m2txt.Text * 2

            Timer1.Enabled = False


        End If

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'yinicial = masaPb.Location.Y
        yinicial = masaPb.Top
        yinicial2 = masaPb2.Top
        altoinicial = resortePb1.Width
    End Sub



End Class
