use chrono::Local;
use clap::{Parser, Subcommand};
use std::fs;
use std::path::PathBuf;

/// Git commit hash embedded at compile time
const GIT_HASH: &str = env!("GIT_HASH");

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
        /// Original document (before changes)
        doc1: PathBuf,

        /// Modified document (after changes)
        doc2: PathBuf,

        /// Output document path (default: redline-DATETIME-COMMIT.docx)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Document type: auto, docx, xlsx, pptx
        #[arg(short = 't', long, default_value = "auto")]
        doc_type: String,

        /// Output revision statistics as JSON
        #[arg(long)]
        json: bool,

        /// Author name for revisions (defaults to doc2's LastModifiedBy or Creator)
        #[arg(long)]
        author: Option<String>,

        /// Date/time for revisions in ISO 8601 format (defaults to doc2's modified date)
        #[arg(long)]
        date: Option<String>,

        /// Trace LCS algorithm for a specific section (e.g., "3.1", "(b)")
        #[arg(long)]
        trace_section: Option<String>,

        /// Trace LCS algorithm for paragraphs starting with this text
        #[arg(long)]
        trace_paragraph: Option<String>,

        /// Output file for LCS trace JSON (default: lcs-trace.json)
        #[arg(long, default_value = "lcs-trace.json")]
        trace_output: PathBuf,

        /// Detail threshold for comparison (0.0-1.0, default: 0.15)
        /// Lower values = more granular word-level changes
        /// Higher values = more coalesced paragraph-level changes
        #[arg(long, default_value = "0.15")]
        detail_threshold: f64,

        /// Transform tracked changes to visual formatting (colored text, no revision markers)
        #[arg(long)]
        visual_redline: bool,

        /// Color for insertions in visual redline mode (hex RGB, e.g., 0000FF for blue)
        #[arg(long, default_value = "0000FF")]
        insertion_color: String,

        /// Color for deletions in visual redline mode (hex RGB, e.g., FF0000 for red)
        #[arg(long, default_value = "FF0000")]
        deletion_color: String,

        /// Color for moved content in visual redline mode (hex RGB, e.g., 008000 for green)
        #[arg(long, default_value = "008000")]
        move_color: String,

        /// Skip summary table in visual redline output
        #[arg(long)]
        no_summary_table: bool,
    },
    /// Count revisions between two documents (without generating output)
    Count {
        /// Original document (before changes)
        doc1: PathBuf,

        /// Modified document (after changes)
        doc2: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
    /// Display information about a document
    Info {
        /// Document to analyze
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

/// Generate default output filename: redline-YYYYMMDD-HHMMSS-COMMIT.docx
fn generate_output_filename(doc_type: &str) -> PathBuf {
    let now = Local::now();
    let datetime = now.format("%Y%m%d-%H%M%S").to_string();
    let extension = match doc_type {
        "docx" => "docx",
        "xlsx" => "xlsx",
        "pptx" => "pptx",
        _ => "docx",
    };
    PathBuf::from(format!("redline-{}-{}.{}", datetime, GIT_HASH, extension))
}

fn main() {
    let cli = Cli::parse();

    let result = match cli.command {
        Commands::Compare {
            doc1,
            doc2,
            output,
            doc_type,
            json,
            author,
            date,
            trace_section,
            trace_paragraph,
            trace_output,
            detail_threshold,
            visual_redline,
            insertion_color,
            deletion_color,
            move_color,
            no_summary_table,
        } => run_compare(
            &doc1,
            &doc2,
            output,
            &doc_type,
            json,
            author,
            date,
            trace_section,
            trace_paragraph,
            trace_output,
            detail_threshold,
            visual_redline,
            &insertion_color,
            &deletion_color,
            &move_color,
            no_summary_table,
        ),

        Commands::Count { doc1, doc2, json } => run_count(&doc1, &doc2, json),

        Commands::Info { file } => run_info(&file),
    };

    if let Err(e) = result {
        eprintln!("Error: {}", e);
        std::process::exit(1);
    }
}

fn run_compare(
    doc1: &PathBuf,
    doc2: &PathBuf,
    output: Option<PathBuf>,
    doc_type: &str,
    json: bool,
    author: Option<String>,
    date: Option<String>,
    trace_section: Option<String>,
    trace_paragraph: Option<String>,
    trace_output: PathBuf,
    detail_threshold: f64,
    visual_redline: bool,
    insertion_color: &str,
    deletion_color: &str,
    move_color: &str,
    no_summary_table: bool,
) -> Result<(), String> {
    let doc_type = detect_doc_type(doc1, doc_type)?;

    // Generate output path if not specified
    let output = output.unwrap_or_else(|| generate_output_filename(doc_type));

    // Read input files
    let bytes1 = fs::read(doc1).map_err(|e| format!("Failed to read {}: {}", doc1.display(), e))?;
    let bytes2 = fs::read(doc2).map_err(|e| format!("Failed to read {}: {}", doc2.display(), e))?;

    match doc_type {
        "docx" => {
            let parsed_doc1 = redline_core::WmlDocument::from_bytes(&bytes1)
                .map_err(|e| format!("Failed to parse {}: {}", doc1.display(), e))?;
            let parsed_doc2 = redline_core::WmlDocument::from_bytes(&bytes2)
                .map_err(|e| format!("Failed to parse {}: {}", doc2.display(), e))?;

            // Extract metadata from doc2 for defaults
            let props = parsed_doc2.package().get_core_properties();

            // Determine author: CLI arg -> lastModifiedBy -> creator -> "Redline"
            let final_author = author
                .or(props.last_modified_by)
                .or(props.creator)
                .unwrap_or_else(|| "Redline".to_string());

            // Determine date: CLI arg -> modified -> None (lets Settings default to current time)
            let final_date = date.or(props.modified);

            // Create settings with resolved author, date, and detail threshold
            let mut settings = redline_core::WmlComparerSettings::default();
            settings.author_for_revisions = Some(final_author);
            if let Some(d) = final_date {
                settings.date_time_for_revisions = Some(d);
            }
            settings.detail_threshold = detail_threshold;

            // Set up LCS trace filter if requested
            #[cfg(feature = "trace")]
            if trace_section.is_some() || trace_paragraph.is_some() {
                settings.lcs_trace_filter = Some(redline_core::LcsTraceFilter {
                    section: trace_section.clone(),
                    paragraph_prefix: trace_paragraph.clone(),
                    output_path: Some(trace_output.clone()),
                });
            }

            // Compare documents
            let result =
                redline_core::WmlComparer::compare(&parsed_doc1, &parsed_doc2, Some(&settings))
                    .map_err(|e| format!("Comparison failed: {}", e))?;

            // Count revisions for reporting
            let insertions = result.insertions;
            let deletions = result.deletions;

            // Apply visual redline transformation if requested
            let final_document = if visual_redline {
                let visual_settings = redline_core::wml::VisualRedlineSettings {
                    insertion_color: insertion_color.to_string(),
                    deletion_color: deletion_color.to_string(),
                    move_color: move_color.to_string(),
                    add_summary_table: !no_summary_table,
                    older_filename: Some(doc1.display().to_string()),
                    newer_filename: Some(doc2.display().to_string()),
                    ..Default::default()
                };

                let compared_doc = redline_core::WmlDocument::from_bytes(&result.document)
                    .map_err(|e| format!("Failed to load comparison result: {}", e))?;

                let visual_result =
                    redline_core::wml::render_visual_redline(&compared_doc, &visual_settings)
                        .map_err(|e| format!("Visual redline transformation failed: {}", e))?;

                eprintln!(
                    "Visual redline: {} insertions, {} deletions, {} moves detected",
                    visual_result.insertions, visual_result.deletions, visual_result.moves
                );

                visual_result.document
            } else {
                result.document
            };

            // Write output
            fs::write(&output, &final_document)
                .map_err(|e| format!("Failed to write {}: {}", output.display(), e))?;

            // Write LCS trace if it was captured
            #[cfg(feature = "trace")]
            if let Some(ref traces) = result.lcs_traces {
                if !traces.is_empty() {
                    let trace_json = serde_json::to_string_pretty(traces)
                        .map_err(|e| format!("Failed to serialize trace: {}", e))?;
                    fs::write(&trace_output, trace_json)
                        .map_err(|e| format!("Failed to write trace file: {}", e))?;
                    eprintln!("LCS trace written to: {}", trace_output.display());
                }
            }

            // Suppress unused variable warnings when trace feature is disabled
            #[cfg(not(feature = "trace"))]
            let _ = (&trace_section, &trace_paragraph, &trace_output);

            // Report results
            if json {
                println!(
                    r#"{{"insertions":{},"deletions":{},"total":{},"output":"{}","commit":"{}"}}"#,
                    insertions,
                    deletions,
                    insertions + deletions,
                    output.display(),
                    GIT_HASH
                );
            } else {
                println!("Comparison complete:");
                println!("  Insertions: {}", insertions);
                println!("  Deletions:  {}", deletions);
                println!("  Total:      {}", insertions + deletions);
                println!("  Output:     {}", output.display());
                println!("  Commit:     {}", GIT_HASH);
            }
        }
        "xlsx" => {
            let parsed_doc1 = redline_core::SmlDocument::from_bytes(&bytes1)
                .map_err(|e| format!("Failed to parse {}: {}", doc1.display(), e))?;
            let parsed_doc2 = redline_core::SmlDocument::from_bytes(&bytes2)
                .map_err(|e| format!("Failed to parse {}: {}", doc2.display(), e))?;

            // Determine author: CLI arg -> "Redline"
            let final_author = author.unwrap_or_else(|| "Redline".to_string());

            let mut settings = redline_core::SmlComparerSettings::default();
            settings.author_for_changes = final_author;
            // SmlComparerSettings doesn't have a generic numeric tolerance exposed via CLI args mapping logic yet
            // But we can map detail_threshold to something if needed, or ignore.
            // SmlComparer uses numeric_tolerance (default 0.0).

            let (marked_doc, result) = redline_core::SmlComparer::compare_and_render(
                &parsed_doc1,
                &parsed_doc2,
                Some(&settings),
            )
            .map_err(|e| format!("Comparison failed: {}", e))?;

            let output_bytes = marked_doc
                .to_bytes()
                .map_err(|e| format!("Failed to serialize output: {}", e))?;

            fs::write(&output, output_bytes)
                .map_err(|e| format!("Failed to write {}: {}", output.display(), e))?;

            if json {
                println!(
                    r#"{{"insertions":{},"deletions":{},"total":{},"output":"{}","commit":"{}"}}"#,
                    result.cells_added(),
                    result.cells_deleted(),
                    result.total_changes(),
                    output.display(),
                    GIT_HASH
                );
            } else {
                println!("Comparison complete:");
                println!("  Cells Added:   {}", result.cells_added());
                println!("  Cells Deleted: {}", result.cells_deleted());
                println!("  Total Changes: {}", result.total_changes());
                println!("  Output:        {}", output.display());
                println!("  Commit:        {}", GIT_HASH);
            }
        }
        "pptx" => {
            let parsed_doc1 = redline_core::PmlDocument::from_bytes(&bytes1)
                .map_err(|e| format!("Failed to parse {}: {}", doc1.display(), e))?;
            let parsed_doc2 = redline_core::PmlDocument::from_bytes(&bytes2)
                .map_err(|e| format!("Failed to parse {}: {}", doc2.display(), e))?;

            let final_author = author.unwrap_or_else(|| "Redline".to_string());

            let mut settings = redline_core::PmlComparerSettings::default();
            settings.author_for_changes = final_author.clone();
            settings.author_for_revisions = final_author;
            // Parse date if provided
            if let Some(d) = date {
                if let Ok(parsed_date) = chrono::DateTime::parse_from_rfc3339(&d) {
                    settings.date_time_for_revisions = parsed_date.with_timezone(&chrono::Utc);
                }
            }

            // Get result for stats
            let result =
                redline_core::PmlComparer::compare(&parsed_doc1, &parsed_doc2, Some(&settings))
                    .map_err(|e| format!("Comparison failed: {}", e))?;

            // Render marked presentation
            let marked_doc =
                redline_core::pml::render_marked_presentation(&parsed_doc2, &result, &settings)
                    .map_err(|e| format!("Rendering failed: {}", e))?;

            let output_bytes = marked_doc
                .to_bytes()
                .map_err(|e| format!("Failed to serialize output: {}", e))?;

            fs::write(&output, output_bytes)
                .map_err(|e| format!("Failed to write {}: {}", output.display(), e))?;

            if json {
                println!(
                    r#"{{"total":{},"output":"{}","commit":"{}"}}"#,
                    result.total_changes,
                    output.display(),
                    GIT_HASH
                );
            } else {
                println!("Comparison complete:");
                println!("  Slides Inserted: {}", result.slides_inserted);
                println!("  Slides Deleted:  {}", result.slides_deleted);
                println!("  Total Changes:   {}", result.total_changes);
                println!("  Output:          {}", output.display());
                println!("  Commit:          {}", GIT_HASH);
            }
        }
        _ => unreachable!(),
    }

    Ok(())
}

fn run_count(doc1: &PathBuf, doc2: &PathBuf, json: bool) -> Result<(), String> {
    let doc_type = detect_doc_type(doc1, "auto")?;

    let bytes1 = fs::read(doc1).map_err(|e| format!("Failed to read {}: {}", doc1.display(), e))?;
    let bytes2 = fs::read(doc2).map_err(|e| format!("Failed to read {}: {}", doc2.display(), e))?;

    match doc_type {
        "docx" => {
            let parsed_doc1 = redline_core::WmlDocument::from_bytes(&bytes1)
                .map_err(|e| format!("Failed to parse {}: {}", doc1.display(), e))?;
            let parsed_doc2 = redline_core::WmlDocument::from_bytes(&bytes2)
                .map_err(|e| format!("Failed to parse {}: {}", doc2.display(), e))?;

            let result = redline_core::WmlComparer::compare(&parsed_doc1, &parsed_doc2, None)
                .map_err(|e| format!("Failed to count revisions: {}", e))?;

            let insertions = result.insertions;
            let deletions = result.deletions;

            if json {
                println!(
                    r#"{{"insertions":{},"deletions":{},"total":{}}}"#,
                    insertions,
                    deletions,
                    insertions + deletions
                );
            } else {
                println!("Revision count:");
                println!("  Insertions: {}", insertions);
                println!("  Deletions:  {}", deletions);
                println!("  Total:      {}", insertions + deletions);
            }
        }
        _ => {
            return Err(format!(
                "Revision counting not implemented for {} files",
                doc_type
            ));
        }
    }

    Ok(())
}

fn run_info(file: &PathBuf) -> Result<(), String> {
    let bytes = fs::read(file).map_err(|e| format!("Failed to read {}: {}", file.display(), e))?;

    let doc_type = detect_doc_type(file, "auto")?;

    println!("Document: {}", file.display());
    println!("Type: {}", doc_type);
    println!("Size: {} bytes", bytes.len());

    Ok(())
}
