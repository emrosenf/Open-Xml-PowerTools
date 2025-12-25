use clap::{Parser, Subcommand};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "redline")]
#[command(about = "OOXML document comparison tool", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    Compare {
        #[arg(short, long)]
        source1: PathBuf,

        #[arg(short = '2', long)]
        source2: PathBuf,

        #[arg(short, long)]
        output: PathBuf,

        #[arg(short = 't', long, default_value = "auto")]
        doc_type: String,

        #[arg(long)]
        json: bool,
    },
    Info {
        #[arg(short, long)]
        file: PathBuf,
    },
}

fn main() {
    let cli = Cli::parse();

    match cli.command {
        Commands::Compare {
            source1,
            source2,
            output,
            doc_type,
            json,
        } => {
            println!("Comparing documents:");
            println!("  Source 1: {}", source1.display());
            println!("  Source 2: {}", source2.display());
            println!("  Output: {}", output.display());
            println!("  Type: {}", doc_type);
            println!("  JSON output: {}", json);
            
            eprintln!("CLI not yet implemented - Phase 6");
            std::process::exit(1);
        }
        Commands::Info { file } => {
            println!("Document info for: {}", file.display());
            
            eprintln!("CLI not yet implemented - Phase 6");
            std::process::exit(1);
        }
    }
}
