using EnvDTE;
using EnvDTE80;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using CodeNamespace = EnvDTE.CodeNamespace;

namespace PointInterruptionConsole
{
    class Program
    {
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        static void Main(string[] args)
        {
            try
            {
                // Obtener la instancia activa de Visual Studio
                DTE2 dte = GetRunningInstance();
                if (dte == null)
                {
                    Console.WriteLine("No se encontró ninguna instancia de Visual Studio.");
                    return;
                }

                // Agregar puntos de interrupción a todos los métodos
                AddBreakpointsToAllMethods(dte);
                Console.WriteLine("Puntos de interrupción añadidos a todos los métodos.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static DTE2 GetRunningInstance()
        {
            try
            {
                CreateBindCtx(0, out var ctx);
                ctx.GetRunningObjectTable(out var rot);
                rot.EnumRunning(out var enumMoniker);
                enumMoniker.Reset();

                var monikers = new IMoniker[1];
                while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    rot.GetObject(monikers[0], out var comObject);
                    var dte = comObject as DTE2;
                    if (dte != null)
                    {
                        return dte;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return null;
        }

        static void AddBreakpointsToAllMethods(DTE2 dte)
        {
            try
            {
                // Obtener los proyectos de la solución
                Projects projects = dte.Solution.Projects;
                foreach (Project project in projects)
                {
                    AddBreakpointsToProject(project);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void AddBreakpointsToProject(Project project)
        {
            try
            {
                // Recorrer los elementos del proyecto
                foreach (ProjectItem item in project.ProjectItems)
                {
                    ProcessProjectItem(item);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void ProcessProjectItem(ProjectItem item)
        {
            try
            {
                // Si el archivo tiene un modelo de código, procesar sus elementos
                if (item.FileCodeModel != null)
                {
                    foreach (CodeElement element in item.FileCodeModel.CodeElements)
                    {
                        if (element.Kind == vsCMElement.vsCMElementNamespace)
                        {
                            CodeNamespace ns = (CodeNamespace)element;
                            ProcessNamespace(ns);
                        }
                    }
                }

                // Recursivamente procesar los subelementos
                if (item.ProjectItems != null)
                {
                    foreach (ProjectItem subItem in item.ProjectItems)
                    {
                        ProcessProjectItem(subItem);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void ProcessNamespace(CodeNamespace ns)
        {
            int retryCount = 5; // Número de reintentos permitidos
            while (retryCount > 0)
            {
                try
                {
                    foreach (CodeElement element in ns.Members)
                    {
                        if (element.Kind == vsCMElement.vsCMElementClass)
                        {
                            CodeClass cls = (CodeClass)element;
                            ProcessClass(cls);
                        }
                    }
                    break; // Si el proceso fue exitoso, salir del bucle
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x8001010A) // Código específico de RPC_E_SERVERCALL_RETRYLATER
                {
                    Console.WriteLine("Visual Studio está ocupado. Reintentando en 500ms...");
                    System.Threading.Thread.Sleep(500); // Esperar medio segundo antes de reintentar
                    retryCount--;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    break;
                }
            }

            if (retryCount == 0)
            {
                Console.WriteLine("Se agotaron los reintentos. Visual Studio sigue ocupado.");
            }
        }

        static void ProcessClass(CodeClass cls)
        {
            try
            {
                foreach (CodeElement element in cls.Members)
                {
                    if (element.Kind == vsCMElement.vsCMElementFunction)
                    {
                        CodeFunction func = (CodeFunction)element;
                        AddBreakpoint(func);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void AddBreakpoint(CodeFunction func)
        {
            try
            {
                TextPoint startPoint = func.GetStartPoint(vsCMPart.vsCMPartBody);
                func.DTE.Debugger.Breakpoints.Add("", func.ProjectItem.FileNames[1], startPoint.Line, 1, "",
                    dbgBreakpointConditionType.dbgBreakpointConditionTypeWhenTrue, "", "", 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
