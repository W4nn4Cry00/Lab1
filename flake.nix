{
  description = "Dev environment for Alina's Lab 1 & SRSP";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixpkgs-unstable";
  };

  outputs = { self, nixpkgs }:
    let
      system = "x86_64-linux";
      pkgs = import nixpkgs { inherit system; };
      pythonEnv = pkgs.python3.withPackages (ps: with ps; [
        python-docx
      ]);
    in
    {
      devShells.${system}.default = pkgs.mkShell {
        packages = with pkgs; [
          pythonEnv
          pandoc
          typst
          tree
        ];
        shellHook = ''
          echo "✨ DevShell active for Alina's project! ✨"
        '';
      };
    };
}
