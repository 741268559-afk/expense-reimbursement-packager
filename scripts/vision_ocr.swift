import Foundation
import Vision
import ImageIO
import CoreGraphics

struct OCRLine: Codable {
    let text: String
    let confidence: Float
    let x: Double
    let y: Double
    let w: Double
    let h: Double
}

struct OCRResult: Codable {
    let file: String
    let path: String
    let lines: [OCRLine]
}

let args = CommandLine.arguments.dropFirst()
guard !args.isEmpty else {
    fputs("Usage: swift vision_ocr.swift image1 image2 ...\n", stderr)
    exit(2)
}

let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.usesLanguageCorrection = true
if #available(macOS 12.0, *) {
    request.recognitionLanguages = ["zh-Hans", "zh-Hant", "en-US"]
}

var results: [OCRResult] = []
for inputPath in args {
    let url = URL(fileURLWithPath: inputPath)
    guard let source = CGImageSourceCreateWithURL(url as CFURL, nil),
          let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
        results.append(OCRResult(file: url.lastPathComponent, path: inputPath, lines: []))
        continue
    }

    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    do {
        try handler.perform([request])
        let observations = request.results ?? []
        let lines = observations.compactMap { obs -> OCRLine? in
            guard let candidate = obs.topCandidates(1).first else { return nil }
            return OCRLine(
                text: candidate.string,
                confidence: candidate.confidence,
                x: Double(obs.boundingBox.origin.x),
                y: Double(obs.boundingBox.origin.y),
                w: Double(obs.boundingBox.size.width),
                h: Double(obs.boundingBox.size.height)
            )
        }.sorted { a, b in
            if abs(a.y - b.y) > 0.012 { return a.y > b.y }
            return a.x < b.x
        }
        results.append(OCRResult(file: url.lastPathComponent, path: inputPath, lines: lines))
    } catch {
        fputs("OCR failed for \(inputPath): \(error)\n", stderr)
        results.append(OCRResult(file: url.lastPathComponent, path: inputPath, lines: []))
    }
}

let encoder = JSONEncoder()
encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
let data = try encoder.encode(results)
FileHandle.standardOutput.write(data)
