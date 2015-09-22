using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace Adil.DAL
{
    /// <summary>
    /// The SQLClient class is intended to encapsulate high performance, scalable best practices for
    /// SQL access
    /// </summary>
    public sealed class SQLClient : IDisposable
    {
        private string _connectionString;

        private SqlConnection _sqlConn;

        #region Constructor

        public SQLClient() : this("Default") { }

        public SQLClient(string key)
        {
            this._connectionString = ConfigurationManager.ConnectionStrings[key].ConnectionString;
            this._sqlConn = new SqlConnection(this._connectionString);
        }

        #endregion

        #region Public Method

        public int ExecuteNonQuery(CommandType commandType, string commandText, SqlTransaction transaction, params SqlParameter[] commandParameters)
        {
            try
            {
                SqlCommand cmd = CreateCommand(commandType, commandText, transaction, commandParameters);

                if (this.OpenConnection())
                {
                    return cmd.ExecuteNonQuery();
                }
                else
                {
                    return -1;
                }
            }
            finally
            {
                this.CloseConnection();
            }
        }

        public int ExecuteNonQuery(CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            try
            {
                SqlCommand cmd = CreateCommand(commandType, commandText, null, commandParameters);

                if (this.OpenConnection())
                {
                    return cmd.ExecuteNonQuery();
                }
                else
                {
                    return -1;
                }
            }
            finally
            {
                this.CloseConnection();
            }
        }

        public int ExecuteNonQuery(string spName, SqlTransaction transaction, params SqlParameter[] commandParameters)
        {
            try
            {
                SqlCommand cmd = CreateCommand(CommandType.StoredProcedure, spName, transaction, commandParameters);

                if (this.OpenConnection())
                {
                    return cmd.ExecuteNonQuery();
                }
                else
                {
                    return -1;
                }

            }
            finally
            {
                this.CloseConnection();
            }
        }

        public int ExecuteNonQuery(string spName, params SqlParameter[] commandParameters)
        {
            try
            {
                SqlCommand cmd = CreateCommand(CommandType.StoredProcedure, spName, null, commandParameters);

                if (this.OpenConnection())
                {
                    return cmd.ExecuteNonQuery();
                }
                else
                {
                    return -1;
                }
            }
            finally
            {
                this.CloseConnection();
            }
        }

        public DataSet ExecuteDataset(CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = CreateCommand(commandType, commandText, null, commandParameters);

            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
        }

        public DataSet ExecuteDataset(string spName, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = CreateCommand(CommandType.StoredProcedure, spName, null, commandParameters);

            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
        }

        public object ExecuteScalar(CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            try
            {

                SqlCommand cmd = CreateCommand(commandType, commandText, null, commandParameters);

                if (this.OpenConnection())
                {
                    return cmd.ExecuteScalar();
                }
                else
                {
                    return null;
                }

            }
            finally
            {
                this.CloseConnection();
            }
        }

        public object ExecuteScalar(string spName, params SqlParameter[] commandParameters)
        {
            try
            {
                SqlCommand cmd = CreateCommand(CommandType.StoredProcedure, spName, null, commandParameters);

                if (this.OpenConnection())
                {
                    return cmd.ExecuteScalar();
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                this.CloseConnection();
            }
        }

        #endregion

        #region Private Method

        private SqlCommand CreateCommand(CommandType commandType, string commandText, SqlTransaction transaction, params SqlParameter[] commandParameters)
        {
            if (commandText == null || commandText.Length == 0) throw new ArgumentNullException("Sql command text");

            SqlCommand cmd = new SqlCommand(commandText, this._sqlConn);

            if (transaction != null)
            {
                if (transaction.Connection == null) throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
                cmd.Transaction = transaction;
            }

            if (commandParameters != null)
            {
                AttachParameters(cmd, commandParameters);
            }

            return cmd;
        }

        private void AttachParameters(SqlCommand command, SqlParameter[] commandParameters)
        {
            if (command == null) throw new ArgumentNullException("SqlCommand reference");
            if (commandParameters != null)
            {
                command.Parameters.Clear();

                foreach (SqlParameter p in commandParameters)
                {
                    if (p != null)
                    {
                        // Check for derived output value with no value assigned
                        if ((p.Direction == ParameterDirection.InputOutput ||
                                p.Direction == ParameterDirection.Input) &&
                                (p.Value == null))
                        {
                            p.Value = DBNull.Value;
                        }
                        command.Parameters.Add(p);
                    }
                }
            }
        }

        private bool OpenConnection()
        {
            try
            {
                if (this._sqlConn.State != ConnectionState.Open) this._sqlConn.Open();

                return true;
            }
            catch (Exception ex)
            {
                //TODO: Need to log error 
                return false;
            }
        }

        private void CloseConnection()
        {
            try
            {
                if (this._sqlConn.State != ConnectionState.Closed) this._sqlConn.Close();
            }
            catch (Exception ex)
            {
                //TODO: Need to log error 
            }
        }



        #endregion

        #region Disposing

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (this._sqlConn != null)
                {
                    this.CloseConnection();
                    this._sqlConn.Dispose();
                    this._sqlConn = null;
                }
            }
        }

        #endregion
    }
}
