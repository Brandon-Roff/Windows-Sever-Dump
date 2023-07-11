#Copy AD Group/User memberships
#Author Brandon Roff

function Copy-ADGroupUserMembership
{
   [CmdletBinding(DefaultParameterSetName = 'default',
      ConfirmImpact = 'Low',
      SupportsShouldProcess)]
   param
   (
      [Parameter(Mandatory,
         ValueFromPipeline,
         ValueFromPipelineByPropertyName,
         Position = 0,
         HelpMessage = 'Source-Group Object.')]
      [ValidateNotNullOrEmpty()]
      [Alias('Source')]
      [string]
      $SourceGroup,
      [Parameter(Mandatory,
         ValueFromPipeline,
         ValueFromPipelineByPropertyName,
         Position = 1,
         HelpMessage = 'Target-Group Object.')]
      [ValidateNotNullOrEmpty()]
      [Alias('Target')]
      [string]
      $TargetGroup,
      [Parameter(ParameterSetName = 'full',
         ValueFromPipeline,
         ValueFromPipelineByPropertyName,
         Position = 2)]
      [Alias('RemoveTargetOnlyMembers')]
      [switch]
      $full = $null,
      [Parameter(ParameterSetName = 'sync',
         ValueFromPipeline,
         ValueFromPipelineByPropertyName,
         Position = 2)]
      [Alias('MakeFullSync')]
      [switch]
      $sync = $null
   )

   begin
   {
      if ($pscmdlet.ShouldProcess('Groups', 'Get information from Active Directory'))
      {
         try
         {
            $SourceMembers = (Get-ADGroupMember -Identity $SourceGroup -ErrorAction Stop | Select-Object -ExpandProperty distinguishedName | Sort-Object)
            $TargetMembers = (Get-ADGroupMember -Identity $TargetGroup -ErrorAction Stop | Select-Object -ExpandProperty distinguishedName | Sort-Object)

            # Check if we have any diferences
            if (($SourceMembers) -and ($TargetMembers))
            {
               # Yep, there are differences
               $Differences = (Compare-Object -ReferenceObject $SourceMembers -DifferenceObject $TargetMembers)
            }
            elseif (($SourceMembers) -and (-not($TargetMembers)))
            {
               # Target has no members
               $Differences = 'SourceOnly'
            }
            elseif (-not($SourceMembers))
            {
               # Source has no members
               Write-Error -Message ('{0} has no members!' -f $SourceGroup) -ErrorAction Stop
            }
            else
            {
               # Nope, there are no differences
               $Differences = $null
            }
         }
         catch
         {
            # get error record
            [Management.Automation.ErrorRecord]$e = $_

            # retrieve information about runtime error
            $info = [PSCustomObject]@{
               Exception = $e.Exception.Message
               Reason    = $e.CategoryInfo.Reason
               Target    = $e.CategoryInfo.TargetName
               Script    = $e.InvocationInfo.ScriptName
               Line      = $e.InvocationInfo.ScriptLineNumber
               Column    = $e.InvocationInfo.OffsetInLine
            }

            $info | Out-String | Write-Verbose

            Write-Error -Message $e.Exception.Message -ErrorAction Stop

            break
         }
      }
   }

   process
   {

      switch ($pscmdlet.ParameterSetName)
      {
         'full'
         {
            if ($pscmdlet.ShouldProcess($TargetGroup, 'Set'))
            {
               if ($Differences)
               {
                  Write-Verbose -Message 'Remove Target-User from all groups where the Source-User is not a member of.'

                  $TargetOnlyMembers = ($Differences | Where-Object -Property SideIndicator -EQ -Value '=>')

                  if ($TargetOnlyMembers)
                  {
                     try
                     {
                        foreach ($TargetOnlyMember in $TargetOnlyMembers.InputObject)
                        {
                           Write-Verbose -Message ('Process: {0}' -f $TargetOnlyMember)

                           $paramRemoveADGroupMember = @{
                              Identity    = $TargetGroup
                              Members     = $TargetOnlyMember
                              ErrorAction = 'Stop'
                              Confirm     = $false
                           }
                           $null = (Remove-ADGroupMember @paramRemoveADGroupMember -Verbose)
                        }
                     }
                     catch
                     {
                        # get error record
                        [Management.Automation.ErrorRecord]$e = $_

                        # retrieve information about runtime error
                        $info = [PSCustomObject]@{
                           Exception = $e.Exception.Message
                           Reason    = $e.CategoryInfo.Reason
                           Target    = $e.CategoryInfo.TargetName
                           Script    = $e.InvocationInfo.ScriptName
                           Line      = $e.InvocationInfo.ScriptLineNumber
                           Column    = $e.InvocationInfo.OffsetInLine
                        }

                        $info | Out-String | Write-Verbose

                        Write-Warning -Message $e.Exception.Message -ErrorAction Continue -WarningAction Continue
                     }
                  }
                  else
                  {
                     Write-Verbose -Message 'No group difference found where the Target-User is a member and Source-User is not.'
                  }
               }
            }
         }
         'sync'
         {
            if ($pscmdlet.ShouldProcess($SourceGroup, 'Set'))
            {
               if ($Differences)
               {
                  Write-Verbose -Message 'Make the Source-user a Member of all Groups only the Target-User is a member of.'

                  $TargetOnlyMembers = ($Differences | Where-Object -Property SideIndicator -EQ -Value '=>')

                  if ($TargetOnlyMembers)
                  {
                     Write-Verbose -Message ('Process: {0}' -f $TargetOnlyMembers)

                     try
                     {
                        $paramAddADGroupMember = @{
                           Identity    = $SourceGroup
                           Members     = $TargetOnlyMembers.InputObject
                           ErrorAction = 'Stop'
                           Confirm     = $false
                        }
                        $null = (Add-ADGroupMember @paramAddADGroupMember)
                     }
                     catch
                     {
                        # get error record
                        [Management.Automation.ErrorRecord]$e = $_

                        # retrieve information about runtime error
                        $info = [PSCustomObject]@{
                           Exception = $e.Exception.Message
                           Reason    = $e.CategoryInfo.Reason
                           Target    = $e.CategoryInfo.TargetName
                           Script    = $e.InvocationInfo.ScriptName
                           Line      = $e.InvocationInfo.ScriptLineNumber
                           Column    = $e.InvocationInfo.OffsetInLine
                        }

                        $info | Out-String | Write-Verbose

                        Write-Warning -Message $e.Exception.Message -ErrorAction Continue -WarningAction Continue
                     }
                  }
                  else
                  {
                     Write-Verbose -Message 'No group difference found where the Target-User is a member and Source-User is not.'
                  }
               }
            }
         }
         'default'
         {
            # Do nothing special
         }
      }

      if ($pscmdlet.ShouldProcess($TargetGroup, 'Set'))
      {
         if ($Differences)
         {
            try
            {
               Write-Verbose -Message 'Process all Source-Group only members.'

               $paramAddADGroupMember = @{
                  Identity    = $TargetGroup
                  ErrorAction = 'Stop'
                  Confirm     = $false
               }

               if ($Differences -eq 'SourceOnly')
               {
                  # Target has no members
                  $paramAddADGroupMember.Members = $SourceMembers
               }
               else
               {
                  $paramAddADGroupMember.Members = ($Differences | Where-Object -Property SideIndicator -EQ -Value '<=' | Select-Object -ExpandProperty InputObject)
               }

               $null = (Add-ADGroupMember @paramAddADGroupMember)
            }
            catch
            {
               # get error record
               [Management.Automation.ErrorRecord]$e = $_

               # retrieve information about runtime error
               $info = [PSCustomObject]@{
                  Exception = $e.Exception.Message
                  Reason    = $e.CategoryInfo.Reason
                  Target    = $e.CategoryInfo.TargetName
                  Script    = $e.InvocationInfo.ScriptName
                  Line      = $e.InvocationInfo.ScriptLineNumber
                  Column    = $e.InvocationInfo.OffsetInLine
               }

               $info | Out-String | Write-Verbose

               Write-Warning -Message $e.Exception.Message -ErrorAction Continue -WarningAction Continue
            }
         }
      }
   }
}

