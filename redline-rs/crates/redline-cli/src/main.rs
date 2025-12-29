use clap::{Parser, Subcommand};
use std::fs;
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "redline")]
#[command(about = "OOXML document comparison tool", long_about = None)]
#[command(version)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Compare two documents and generate a redlined output
    Compare {
        /// First (older/original) document
        #[arg(short = '1', long)]
        source1: PathBuf,

        /// Second (newer/modified) document
        #[arg(short = '2', long)]
        source2: PathBuf,

        /// Output document path
        #[arg(short, long)]
        output: PathBuf,

        /// Document type: auto, docx, xlsx, pptx
        #[arg(short = 't', long, default_value = "auto")]
        doc_type: String,

        /// Output revision statistics as JSON
        #[arg(long)]
        json: bool,

        /// Author name for revisions (defaults to source2's LastModifiedBy or Creator)
        #[arg(long)]
        author: Option<String>,

        /// Date/time for revisions in ISO 8601 format (defaults to source2's modified date)
        #[arg(long)]
        date: Option<String>,
    },
    /// Count revisions between two documents (without generating output)
    Count {
        /// First (older/original) document
        #[arg(short = '1', long)]
        source1: PathBuf,

        /// Second (newer/modified) document
        #[arg(short = '2', long)]
        source2: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
    /// Display information about a document
    Info {
        /// Document to analyze
        #[arg(short, long)]
        file: PathBuf,
    },
}

fn detect_doc_type(path: &PathBuf, hint: &str) -> Result<&'static str, String> {
    if hint != "auto" {
        return match hint {
            "docx" => Ok("docx"),
            "xlsx" => Ok("xlsx"),
            "pptx" => Ok("pptx"),
            _ => Err(format!("Unknown document type: {}", hint)),
        };
    }
    
    // Auto-detect from extension
    match path.extension().and_then(|e| e.to_str()) {
        Some("docx") => Ok("docx"),
        Some("xlsx") => Ok("xlsx"),
        Some("pptx") => Ok("pptx"),
        Some(ext) => Err(format!("Unknown file extension: .{}", ext)),
        None => Err("Cannot detect document type without file extension".to_string()),
    }
}

fn main() {
    let cli = Cli::parse();

    let result = match cli.command {
        Commands::Compare {
            source1,
            source2,
            output,
            doc_type,
            json,
            author,
            date,
        } => run_compare(&source1, &source2, &output, &doc_type, json, author, date),
        
        Commands::Count {
            source1,
            source2,
            json,
        } => run_count(&source1, &source2, json),
        
        Commands::Info { file } => run_info(&file),
    };

    if let Err(e) = result {
        eprintln!("Error: {}", e);
        std::process::exit(1);
    }
}

fn run_compare(
    source1: &PathBuf,
    source2: &PathBuf,
    output: &PathBuf,
    doc_type: &str,
    json: bool,
    author: Option<String>,
    date: Option<String>,
) -> Result<(), String> {
    let doc_type = detect_doc_type(source1, doc_type)?;
    
    // Read input files
    let bytes1 = fs::read(source1)
        .map_err(|e| format!("Failed to read {}: {}", source1.display(), e))?;
    let bytes2 = fs::read(source2)
        .map_err(|e| format!("Failed to read {}: {}", source2.display(), e))?;

    match doc_type {
        "docx" => {
            let doc1 = redline_core::WmlDocument::from_bytes(&bytes1)
                .map_err(|e| format!("Failed to parse {}: {}", source1.display(), e))?;
            let doc2 = redline_core::WmlDocument::from_bytes(&bytes2)
                .map_err(|e| format!("Failed to parse {}: {}", source2.display(), e))?;

            // Extract metadata from doc2 for defaults
            let props = doc2.package().get_core_properties();
            
            // Determine author: CLI arg -> lastModifiedBy -> creator -> "Redline"
            let final_author = author
                .or(props.last_modified_by)
                .or(props.creator)
                .unwrap_or_else(|| "Redline".to_string());
                
            // Determine date: CLI arg -> modified -> None (lets Settings default to current time)
            let final_date = date.or(props.modified);

            // Create settings with resolved author and date
            let mut settings = redline_core::WmlComparerSettings::default();
            settings.author_for_revisions = Some(final_author);
            if let Some(d) = final_date {
                settings.date_time_for_revisions = Some(d);
            }

            // Compare documents
            let result = redline_core::WmlComparer::compare(&doc1, &doc2, Some(&settings))
                .map_err(|e| format!("Comparison failed: {}", e))?;

            // Count revisions for reporting
            let insertions = result.insertions;
            let deletions = result.deletions;

            // Write output
            fs::write(output, &result.document)
                .map_err(|e| format!("Failed to write {}: {}", output.display(), e))?;

            // Report results
            if json {
                println!(r#"{{"insertions":{},"deletions":{},"total":{},"output":"{}"}}"#,
                    insertions, deletions, insertions + deletions, output.display());
            } else {
                println!("Comparison complete:");
                println!("  Insertions: {}", insertions);
                println!("  Deletions:  {}", deletions);
                println!("  Total:      {}", insertions + deletions);
                println!("  Output:     {}", output.display());
            }
        }
        "xlsx" => {
            return Err("Excel comparison not yet implemented".to_string());
        }
        "pptx" => {
            return Err("PowerPoint comparison not yet implemented".to_string());
        }
        _ => unreachable!(),
    }

    Ok(())
}

fn run_count(source1: &PathBuf, source2: &PathBuf, json: bool) -> Result<(), String> {
    let doc_type = detect_doc_type(source1, "auto")?;
    
    let bytes1 = fs::read(source1)
        .map_err(|e| format!("Failed to read {}: {}", source1.display(), e))?;
    let bytes2 = fs::read(source2)
        .map_err(|e| format!("Failed to read {}: {}", source2.display(), e))?;

    match doc_type {
        "docx" => {
            let doc1 = redline_core::WmlDocument::from_bytes(&bytes1)
                .map_err(|e| format!("Failed to parse {}: {}", source1.display(), e))?;
            let doc2 = redline_core::WmlDocument::from_bytes(&bytes2)
                .map_err(|e| format!("Failed to parse {}: {}", source2.display(), e))?;

            let result = redline_core::WmlComparer::compare(&doc1, &doc2, None)
                .map_err(|e| format!("Failed to count revisions: {}", e))?;
            
            let insertions = result.insertions;
            let deletions = result.deletions;

            if json {
                println!(r#"{{"insertions":{},"deletions":{},"total":{}}}"#,
                    insertions, deletions, insertions + deletions);
            } else {
                println!("Revision count:");
                println!("  Insertions: {}", insertions);
                println!("  Deletions:  {}", deletions);
                println!("  Total:      {}", insertions + deletions);
            }
        }
        _ => {
            return Err(format!("Revision counting not implemented for {} files", doc_type));
        }
    }

    Ok(())
}

fn run_info(file: &PathBuf) -> Result<(), String> {
    let bytes = fs::read(file)
        .map_err(|e| format!("Failed to read {}: {}", file.display(), e))?;
    
    let doc_type = detect_doc_type(file, "auto")?;
    
    println!("Document: {}", file.display());
    println!("Type: {}", doc_type);
    println!("Size: {} bytes", bytes.len());
    
    // TODO: Add more document info (page count, word count, etc.)
    
    Ok(())
}