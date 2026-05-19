/*
 ClashDetectionCommand detects geometric intersections (clashes) between bodies in the active document.
 Uses a two-stage approach: AABB broad-phase filtering followed by Body.GetCollision() narrow-phase.
 Clashing bodies are selected and results are reported in the status bar and log.
*/

using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Localization;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System;
using System.Collections.Generic;
using System.Linq;
using Application = SpaceClaim.Api.V242.Application;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct2026.ClashDetection
{
    public static class ClashDetectionCommand
    {
        public static void DetectClashes(Window window)
        {
            try
            {
                if (window?.Document == null)
                {
                    L.Status("ClashDetection_Msg_NoDocument", StatusMessageType.Warning);
                    return;
                }

                var mainPart = window.Document.MainPart;
                var allBodies = mainPart.GetDescendants<IDesignBody>().ToList();

                if (allBodies.Count < 2)
                {
                    L.Status("ClashDetection_Msg_Need2Bodies", StatusMessageType.Information);
                    return;
                }

                Logger.Log($"[ClashDetection] Starting clash detection on {allBodies.Count} bodies...");

                // Build world-space bounding boxes for broad phase
                var entries = new List<BodyEntry>();
                foreach (var idb in allBodies)
                {
                    var master = idb.Master;
                    if (master?.Shape == null) continue;

                    try
                    {
                        // Get the transform that maps from occurrence space to master (local) space.
                        // To get world-space bounding box, we compute the box in identity (world) space
                        // by using the inverse of TransformToMaster.
                        var toMaster = idb.TransformToMaster;
                        var toWorld = toMaster.Inverse;
                        var box = master.Shape.GetBoundingBox(toWorld, true);
                        entries.Add(new BodyEntry(idb, master, box));
                    }
                    catch
                    {
                        // Skip bodies whose bounding box cannot be computed
                    }
                }

                if (entries.Count < 2)
                {
                    L.Status("ClashDetection_Msg_NotEnoughValid", StatusMessageType.Information);
                    return;
                }

                // Broad phase: AABB overlap check
                var candidates = new List<(BodyEntry A, BodyEntry B)>();
                for (int i = 0; i < entries.Count; i++)
                {
                    for (int j = i + 1; j < entries.Count; j++)
                    {
                        var boxA = entries[i].WorldBox;
                        var boxB = entries[j].WorldBox;

                        if (boxA.MinCorner.X <= boxB.MaxCorner.X && boxA.MaxCorner.X >= boxB.MinCorner.X &&
                            boxA.MinCorner.Y <= boxB.MaxCorner.Y && boxA.MaxCorner.Y >= boxB.MinCorner.Y &&
                            boxA.MinCorner.Z <= boxB.MaxCorner.Z && boxA.MaxCorner.Z >= boxB.MinCorner.Z)
                        {
                            candidates.Add((entries[i], entries[j]));
                        }
                    }
                }

                Logger.Log($"[ClashDetection] Broad phase: {candidates.Count} AABB-overlapping pairs from {entries.Count} bodies.");

                // Narrow phase: Body.GetCollision()
                var clashes = new List<(IDesignBody A, IDesignBody B)>();
                foreach (var (a, b) in candidates)
                {
                    try
                    {
                        // Both bodies need to be tested in the same coordinate space.
                        // Copy body A's shape into world space, copy body B's shape into world space.
                        using var shapeA = a.DesignBody.Master.Shape.Copy();
                        shapeA.Transform(a.DesignBody.TransformToMaster.Inverse);

                        using var shapeB = b.DesignBody.Master.Shape.Copy();
                        shapeB.Transform(b.DesignBody.TransformToMaster.Inverse);

                        var collision = shapeA.GetCollision(shapeB);
                        if (collision == Collision.Intersect)
                        {
                            clashes.Add((a.DesignBody, b.DesignBody));
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[ClashDetection] Collision check failed for a pair: {ex.Message}");
                    }
                }

                Logger.Log($"[ClashDetection] Narrow phase: {clashes.Count} clashing pairs found.");

                // Highlight clashing bodies
                if (clashes.Count > 0)
                {
                    var clashingBodies = new HashSet<IDesignBody>();
                    foreach (var (a, b) in clashes)
                    {
                        clashingBodies.Add(a);
                        clashingBodies.Add(b);
                        Logger.Log($"[ClashDetection] Clash: \"{a.Master?.Name}\" <-> \"{b.Master?.Name}\"");
                    }

                    // Select clashing bodies so SpaceClaim highlights them
                    window.ActiveContext.Selection = clashingBodies.Cast<IDocObject>().ToList();

                    L.Status("ClashDetection_Msg_Found", StatusMessageType.Warning,
                        clashes.Count, clashingBodies.Count);
                }
                else
                {
                    L.Status("ClashDetection_Msg_None", StatusMessageType.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[ClashDetection] DetectClashes failed: {ex}");
                L.Status("ClashDetection_Err_Failed", StatusMessageType.Error, ex.Message);
            }
        }

        private struct BodyEntry
        {
            public IDesignBody DesignBody;
            public DesignBody Master;
            public Box WorldBox;

            public BodyEntry(IDesignBody designBody, DesignBody master, Box worldBox)
            {
                DesignBody = designBody;
                Master = master;
                WorldBox = worldBox;
            }
        }
    }
}
