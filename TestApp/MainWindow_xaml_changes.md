# MainWindow.xaml — Three exact changes needed

## Change 1: Add a RowDefinition to the right-side runner Grid
**Location:** Line 1911–1916 (inside `<!-- ══ RIGHT: Runner ══ -->`)

Replace:
```xml
<Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>  <!-- script file -->
    <RowDefinition Height="Auto"/>  <!-- runtime + args -->
    <RowDefinition Height="Auto"/>  <!-- working dir + env vars -->
    <RowDefinition Height="*"/>     <!-- log panel -->
    <RowDefinition Height="Auto"/>  <!-- run button -->
</Grid.RowDefinitions>
```

With:
```xml
<Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>  <!-- script file -->
    <RowDefinition Height="Auto"/>  <!-- smart params panel  ← NEW -->
    <RowDefinition Height="Auto"/>  <!-- runtime + args -->
    <RowDefinition Height="Auto"/>  <!-- working dir + env vars -->
    <RowDefinition Height="*"/>     <!-- log panel -->
    <RowDefinition Height="Auto"/>  <!-- run button -->
</Grid.RowDefinitions>
```

---

## Change 2: Bump existing Grid.Row numbers by 1 and name the Runtime+Args border

**Location:** The `<!-- Runtime + Args -->` border (~line 1977).

Change `Grid.Row="1"` → `Grid.Row="2"` AND add `x:Name="ScriptManualArgsContainer"`:

```xml
<!-- Runtime + Args -->
<Border x:Name="ScriptManualArgsContainer"
        Grid.Row="2" Background="#0D1020" BorderBrush="#1A2030"
```

Then bump the remaining three borders:
- `<!-- Working dir -->` border: `Grid.Row="2"` → `Grid.Row="3"`
- `<!-- Log panel -->` border (`x:Name="ScriptLogPanel"`): `Grid.Row="3"` → `Grid.Row="4"`
- `<!-- Run button -->` Grid: `Grid.Row="4"` → `Grid.Row="5"`

---

## Change 3: Insert the smart params panel as Grid.Row="1"

**Location:** Between the closing `</Border>` of the Script File selector and the `<!-- Runtime + Args -->` comment.

Insert this entire block there:

```xml
<!-- Smart params panel -->
<Border x:Name="ScriptDynParamsContainer"
        Grid.Row="1"
        Background="#0D1020"
        BorderBrush="#1A2030"
        BorderThickness="1"
        CornerRadius="8"
        Padding="14,12"
        Margin="0,0,0,10"
        Visibility="Collapsed">
    <StackPanel>
        <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Horizontal">
                <TextBlock Text="SCRIPT PARAMETERS"
                           FontSize="10" FontWeight="Bold" Foreground="#6B7A99"
                           FontFamily="Segoe UI Variable, Segoe UI"
                           VerticalAlignment="Center"/>
                <Border x:Name="ParamDetectionBadge"
                        Background="#1A2C1A" BorderBrush="#2D4A2D"
                        BorderThickness="1" CornerRadius="4"
                        Padding="8,2" Margin="10,0,0,0">
                    <TextBlock x:Name="ParamDetectionLabel"
                               Text="auto-detected"
                               Foreground="#4ADE80" FontSize="10"
                               FontFamily="Segoe UI Variable, Segoe UI"/>
                </Border>
            </StackPanel>
            <Button x:Name="ParamManualToggleBtn"
                    Grid.Column="1"
                    Content="Edit manually instead"
                    Background="Transparent"
                    Foreground="#4A5F88"
                    BorderThickness="0"
                    FontSize="10.5"
                    Cursor="Hand"
                    Click="ParamManualToggle_Click">
                <Button.Style>
                    <Style TargetType="Button">
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <TextBlock x:Name="Tb"
                                               Text="{TemplateBinding Content}"
                                               Foreground="{TemplateBinding Foreground}"
                                               FontSize="{TemplateBinding FontSize}"
                                               TextDecorations="Underline"
                                               Cursor="Hand"/>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="Tb" Property="Foreground" Value="#60A5FA"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Button.Style>
            </Button>
        </Grid>
        <StackPanel x:Name="ScriptDynParamsHost"/>
        <TextBlock Text="  * required"
                   Foreground="#4A5F88" FontSize="10.5"
                   FontFamily="Segoe UI Variable, Segoe UI"
                   Margin="0,8,0,0"/>
    </StackPanel>
</Border>
```
