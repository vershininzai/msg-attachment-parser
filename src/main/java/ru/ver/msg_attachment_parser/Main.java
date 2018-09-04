package ru.ver.msg_attachment_parser;

import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.LinkedList;
import java.util.List;

public class Main {
    public static void main(String... args) throws IOException {
        String rootFolder = "D:/path/to/files";
        Path path = Paths.get(rootFolder);
        final List<Path> files = new LinkedList<>();
        Files.walkFileTree(path, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                if (!attrs.isDirectory() && file.getFileName().toString().endsWith(".msg")) {
                    files.add(file);
                }
                return FileVisitResult.CONTINUE;
            }
        });
        for (Path file : files) {
            try {
                String dir = file.getParent().toString();
                MAPIMessage msg = new MAPIMessage(file.toFile());
                AttachmentChunks[] attachmentMap = msg.getAttachmentFiles();
                if (attachmentMap.length > 0) {
                    for (AttachmentChunks attachmentChunks : attachmentMap) {
                        String attPath = dir + "\\" + file.getFileName().toString() + "_" + attachmentChunks.getAttachLongFileName().toString();
                        try (FileOutputStream fos = new FileOutputStream(attPath)) {
                            fos.write(attachmentChunks.getEmbeddedAttachmentObject());
                        }
                    }
                } else {
                    System.out.println("No attachment");
                }
                System.out.println(file.toFile().getAbsoluteFile() + "  OK");
            } catch (Exception e) {
                System.out.println(file.toFile().getAbsoluteFile() + "  ERROR");
            }
        }
    }
}
