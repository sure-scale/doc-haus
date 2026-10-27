import path from "node:path"

// A fixed current-user ACL operation shared by native checkout and server
// transport storage. Paths are data in the environment, never PowerShell code.
export const windowsPrivateFileScript = [
  "$ErrorActionPreference='Stop'",
  "$target=$env:DOCHAUS_WORD_CACHE_PATH",
  "$user=[System.Security.Principal.WindowsIdentity]::GetCurrent().User",
  "if($env:DOCHAUS_WORD_CACHE_NEW -eq '1'){",
  "  if(Test-Path -LiteralPath $target -PathType Container){",
  "    $acl=[System.Security.AccessControl.DirectorySecurity]::new()",
  "    $rule=[System.Security.AccessControl.FileSystemAccessRule]::new($user,'FullControl','ContainerInherit,ObjectInherit','None','Allow')",
  "  }else{",
  "    $acl=[System.Security.AccessControl.FileSecurity]::new()",
  "    $rule=[System.Security.AccessControl.FileSystemAccessRule]::new($user,'FullControl','Allow')",
  "  }",
  "  $acl.SetOwner($user)",
  "  $acl.SetAccessRuleProtection($true,$false)",
  "  $acl.AddAccessRule($rule)",
  "  Set-Acl -LiteralPath $target -AclObject $acl",
  "}",
  "$acl=Get-Acl -LiteralPath $target",
  "if($acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value -ne $user.Value){throw 'owner'}",
  "$allowed=@($user.Value,'S-1-5-18','S-1-5-32-544')",
  "$own=$false",
  "foreach($rule in $acl.GetAccessRules($true,$true,[System.Security.Principal.SecurityIdentifier])){",
  "  if($rule.AccessControlType -eq 'Allow'){",
  "    if($allowed -notcontains $rule.IdentityReference.Value){throw 'access'}",
  "    if($rule.IdentityReference.Value -eq $user.Value -and ($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute)){ $own=$true }",
  "  }",
  "}",
  "if(-not $own){throw 'access'}",
].join("\n")

export function windowsPrivateFileCommand(file: string, created: boolean, environment: NodeJS.ProcessEnv = process.env) {
  return {
    command: [path.win32.join(environment.SystemRoot || "C:\\Windows", "System32/WindowsPowerShell/v1.0/powershell.exe"),
      "-NoProfile", "-NonInteractive", "-Command", windowsPrivateFileScript],
    env: { ...environment, DOCHAUS_WORD_CACHE_PATH: file, DOCHAUS_WORD_CACHE_NEW: created ? "1" : "0" },
  }
}
